#nullable enable
namespace Divit.Pdf;

using Divit.Packaging;
using Divit.Properties;
using System;
using System.IO;
using System.IO.Compression;
using System.Text;

/// <summary>
/// An image as a PDF image object: its dictionary entries, its data and, for transparency, a soft mask. JPEG goes
/// in as it is; PNG without an alpha channel keeps its compressed data, which PDF reads through the PNG predictors;
/// only PNG with transparency is decoded, to split colour from alpha.
/// </summary>
internal sealed class PdfImage
{
    private PdfImage(int width, int height, string entries, byte[] data, bool compress, PdfImage? mask = null)
    {
        Width = width;
        Height = height;
        Entries = entries;
        Data = data;
        Compress = compress;
        Mask = mask;
    }

    public int Width { get; }

    public int Height { get; }

    /// <summary>
    /// Dictionary entries after the type, size and mask: colour space, bits, filter.
    /// </summary>
    public string Entries { get; }

    public byte[] Data { get; }

    /// <summary>
    /// Whether the data still has to be compressed; JPEG and PNG data already are.
    /// </summary>
    public bool Compress { get; }

    /// <summary>
    /// The transparency, as a grey image of the same size.
    /// </summary>
    public PdfImage? Mask { get; }

    public static PdfImage From(byte[] bytes)
    {
        var info = ImageInfo.Read(bytes);
        if (info == null)
        {
            throw new ArgumentException(Messages.Get(Strings.ImageFormatUnknown));
        }

        try
        {
            if (info.Value.Extension == "jpeg")
            {
                return Jpeg(bytes, info.Value);
            }

            if (info.Value.Extension == "png")
            {
                return Png(bytes);
            }
        }
        catch (IndexOutOfRangeException)
        {
            throw new ArgumentException(Messages.Get(Strings.ImageFormatUnknown));
        }

        throw new NotSupportedException(Messages.Get(Strings.PdfImageUnsupported));
    }

    private static PdfImage Jpeg(byte[] bytes, ImageInfo info)
    {
        // The number of colour components, from the frame header; Adobe's marker means inverted CMYK.
        var components = 3;
        var adobe = false;
        for (var i = 2; i + 9 < bytes.Length;)
        {
            if (bytes[i] != 0xFF)
            {
                break;
            }

            var marker = bytes[i + 1];
            if (marker == 0xFF)
            {
                i++;
                continue;
            }

            var length = (bytes[i + 2] << 8) | bytes[i + 3];
            if (marker == 0xEE && length >= 7 && Encoding.ASCII.GetString(bytes, i + 4, 5) == "Adobe")
            {
                adobe = true;
            }

            if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
            {
                components = bytes[i + 9];
                break;
            }

            i += 2 + length;
        }

        var space = components == 1 ? "/DeviceGray" : components == 4 ? "/DeviceCMYK" : "/DeviceRGB";
        var decode = components == 4 && adobe ? "/Decode [1 0 1 0 1 0 1 0] " : string.Empty;
        return new PdfImage(info.Width, info.Height, "/ColorSpace " + space + " /BitsPerComponent 8 " + decode + "/Filter /DCTDecode ", bytes, false);
    }

    private static PdfImage Png(byte[] bytes)
    {
        int width = 0, height = 0, depth = 8, type = 2, interlace = 0;
        byte[]? palette = null, transparency = null;
        var data = new MemoryStream();
        for (var p = 8; p + 8 <= bytes.Length;)
        {
            var length = BigEndian(bytes, p);
            var chunk = Encoding.ASCII.GetString(bytes, p + 4, 4);
            var body = p + 8;
            switch (chunk)
            {
                case "IHDR":
                    width = BigEndian(bytes, body);
                    height = BigEndian(bytes, body + 4);
                    depth = bytes[body + 8];
                    type = bytes[body + 9];
                    interlace = bytes[body + 12];
                    break;
                case "PLTE":
                    palette = Slice(bytes, body, length);
                    break;
                case "tRNS":
                    transparency = Slice(bytes, body, length);
                    break;
                case "IDAT":
                    data.Write(bytes, body, length);
                    break;
            }

            if (chunk == "IEND")
            {
                break;
            }

            p = body + length + 4;
        }

        if (interlace != 0 || width <= 0 || height <= 0)
        {
            throw new NotSupportedException(Messages.Get(Strings.PdfImageUnsupported));
        }

        var colors = type == 2 ? 3 : type == 6 ? 4 : type == 4 ? 2 : 1;
        var compressed = data.ToArray();
        if (type == 0 || type == 2 || type == 3)
        {
            // The compressed rows as they are, read through the PNG predictors.
            var space = type == 0 ? "/DeviceGray" : type == 2 ? "/DeviceRGB" : Indexed(palette);
            var entries = new StringBuilder("/ColorSpace ").Append(space).Append(" /BitsPerComponent ").Append(Int(depth))
                .Append(" /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors ").Append(Int(colors))
                .Append(" /BitsPerComponent ").Append(Int(depth)).Append(" /Columns ").Append(Int(width)).Append(" >> ");
            PdfImage? mask = null;
            if (transparency != null && type != 3)
            {
                // A colour key: pixels of exactly this grey or colour are left out.
                entries.Append("/Mask [");
                for (var i = 0; i + 1 < transparency.Length && i < 2 * colors; i += 2)
                {
                    var value = Int((transparency[i] << 8) | transparency[i + 1]);
                    entries.Append(value).Append(' ').Append(value).Append(' ');
                }

                entries.Append("] ");
            }
            else if (transparency != null)
            {
                // Alpha per palette entry: a soft mask from the decoded indices.
                var rows = Unfilter(compressed, width, height, colors, depth);
                var stride = ((width * depth) + 7) / 8;
                var alpha = new byte[width * height];
                for (var y = 0; y < height; y++)
                {
                    for (var x = 0; x < width; x++)
                    {
                        var bit = x * depth;
                        var index = (rows[(y * stride) + (bit / 8)] >> (8 - depth - (bit % 8))) & ((1 << depth) - 1);
                        alpha[(y * width) + x] = index < transparency.Length ? transparency[index] : (byte)255;
                    }
                }

                mask = Gray(width, height, alpha);
            }

            return new PdfImage(width, height, entries.ToString(), compressed, false, mask);
        }

        // Grey or colour with alpha: decoded, the colour and the alpha written apart (16-bit samples cut to 8).
        var pixels = Unfilter(compressed, width, height, colors, depth);
        var size = depth / 8;
        var channels = colors - 1;
        var color = new byte[width * height * channels];
        var opacity = new byte[width * height];
        for (int i = 0, o = 0, a = 0; i < pixels.Length; i += colors * size)
        {
            for (var c = 0; c < channels; c++)
            {
                color[o++] = pixels[i + (c * size)];
            }

            opacity[a++] = pixels[i + (channels * size)];
        }

        var colorSpace = channels == 1 ? "/DeviceGray" : "/DeviceRGB";
        return new PdfImage(width, height, "/ColorSpace " + colorSpace + " /BitsPerComponent 8 ", color, true, Gray(width, height, opacity));
    }

    private static PdfImage Gray(int width, int height, byte[] values) =>
        new PdfImage(width, height, "/ColorSpace /DeviceGray /BitsPerComponent 8 ", values, true);

    private static string Indexed(byte[]? palette)
    {
        palette ??= new byte[3];
        var hex = new StringBuilder("[/Indexed /DeviceRGB ").Append(Int((palette.Length / 3) - 1)).Append(" <");
        foreach (var b in palette)
        {
            hex.Append(b.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
        }

        return hex.Append(">]").ToString();
    }

    /// <summary>
    /// The raw rows of a PNG: inflated, each row's filter (none, sub, up, average, Paeth) undone, filter bytes gone.
    /// </summary>
    private static byte[] Unfilter(byte[] compressed, int width, int height, int colors, int depth)
    {
        var stride = ((width * colors * depth) + 7) / 8;
        var step = Math.Max(1, colors * depth / 8);
        var raw = new byte[(stride + 1) * height];
        using (var inflate = new DeflateStream(new MemoryStream(compressed, 2, compressed.Length - 2), CompressionMode.Decompress))
        {
            var read = 0;
            while (read < raw.Length)
            {
                var n = inflate.Read(raw, read, raw.Length - read);
                if (n == 0)
                {
                    break;
                }

                read += n;
            }
        }

        var rows = new byte[stride * height];
        for (var y = 0; y < height; y++)
        {
            var filter = raw[y * (stride + 1)];
            var source = (y * (stride + 1)) + 1;
            var row = y * stride;
            var previous = row - stride;
            for (var x = 0; x < stride; x++)
            {
                int left = x >= step ? rows[row + x - step] : 0;
                int up = y > 0 ? rows[previous + x] : 0;
                int corner = y > 0 && x >= step ? rows[previous + x - step] : 0;
                int value = raw[source + x];
                switch (filter)
                {
                    case 1: value += left; break;
                    case 2: value += up; break;
                    case 3: value += (left + up) / 2; break;
                    case 4:
                        int estimate = left + up - corner, pa = Math.Abs(estimate - left), pb = Math.Abs(estimate - up), pc = Math.Abs(estimate - corner);
                        value += pa <= pb && pa <= pc ? left : pb <= pc ? up : corner;
                        break;
                }

                rows[row + x] = (byte)value;
            }
        }

        return rows;
    }

    private static int BigEndian(byte[] data, int offset) => (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];

    private static byte[] Slice(byte[] data, int offset, int length)
    {
        var slice = new byte[length];
        Buffer.BlockCopy(data, offset, slice, 0, length);
        return slice;
    }

    private static string Int(int value) => OfficeText.Int(value);
}
