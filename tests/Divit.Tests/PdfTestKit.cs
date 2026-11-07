namespace Divit.Tests;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

/// <summary>
/// Checks of saved PDF files, as the Open XML SDK validator checks Office files: the cross-reference table points
/// at every object, and a PDF reader (PdfPig, test-only) opens every page, its fonts and images. Also small images
/// for the tests.
/// </summary>
internal static class PdfTestKit
{
    /// <summary>
    /// A 24 × 16 JPEG, half blue and half orange, drawn with System.Drawing.
    /// </summary>
    internal static readonly byte[] Jpeg = Convert.FromBase64String("/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAQABgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDhK7TwB/zEP+2f/s1cXXaeAP8AmIf9s/8A2ap4k/5FdX/t3/0pGWT/AO+w+f5M7Siiivyo+5P/2Q==");

    /// <summary>
    /// Problems of a PDF file as readable lines; empty when it is sound.
    /// </summary>
    internal static List<string> Validate(string path)
    {
        var problems = new List<string>();
        var bytes = File.ReadAllBytes(path);
        var text = Encoding.ASCII.GetString(bytes);
        if (!text.StartsWith("%PDF-1.7\n", StringComparison.Ordinal))
        {
            problems.Add("no PDF header");
        }

        var startxref = text.LastIndexOf("startxref\n", StringComparison.Ordinal);
        if (startxref < 0 || !text.EndsWith("%%EOF\n", StringComparison.Ordinal))
        {
            problems.Add("no trailer");
            return problems;
        }

        var xref = int.Parse(text.Substring(startxref + 10, text.IndexOf('\n', startxref + 10) - startxref - 10), CultureInfo.InvariantCulture);
        if (string.CompareOrdinal(text, xref, "xref\n0 ", 0, 7) != 0)
        {
            problems.Add("startxref does not point at the table");
            return problems;
        }

        var header = text.Substring(xref + 7, text.IndexOf('\n', xref + 7) - xref - 7);
        var count = int.Parse(header, CultureInfo.InvariantCulture);
        var entry = text.IndexOf('\n', xref + 7) + 1 + 20;
        for (var id = 1; id < count; id++, entry += 20)
        {
            var offset = int.Parse(text.Substring(entry, 10), CultureInfo.InvariantCulture);
            var expected = id.ToString(CultureInfo.InvariantCulture) + " 0 obj";
            if (string.CompareOrdinal(text, offset, expected, 0, expected.Length) != 0)
            {
                problems.Add("object " + expected + " is not at its offset");
            }
        }

        try
        {
            using (var document = UglyToad.PdfPig.PdfDocument.Open(path))
            {
                foreach (var page in document.GetPages())
                {
                    _ = page.Letters.Count;
                    foreach (var image in page.GetImages())
                    {
                        _ = image.RawBytes.Length;
                    }
                }
            }
        }
        catch (Exception e)
        {
            problems.Add("reader: " + e.Message);
        }

        return problems;
    }

    /// <summary>
    /// A PNG of any colour type at 8 bits; row y uses filter y mod 5 (none, sub, up, average, Paeth), so every
    /// predictor is exercised.
    /// </summary>
    internal static byte[] Png(int width, int height, int colorType, Func<int, int, byte[]> pixel, byte[]? palette = null, byte[]? transparency = null)
    {
        var channels = colorType == 2 ? 3 : colorType == 6 ? 4 : colorType == 4 ? 2 : 1;
        var stride = width * channels;
        var rows = new byte[height][];
        for (var y = 0; y < height; y++)
        {
            rows[y] = new byte[stride];
            for (var x = 0; x < width; x++)
            {
                Array.Copy(pixel(x, y), 0, rows[y], x * channels, channels);
            }
        }

        var raw = new MemoryStream();
        for (var y = 0; y < height; y++)
        {
            var filter = y % 5;
            raw.WriteByte((byte)filter);
            for (var i = 0; i < stride; i++)
            {
                int left = i >= channels ? rows[y][i - channels] : 0, up = y > 0 ? rows[y - 1][i] : 0, corner = y > 0 && i >= channels ? rows[y - 1][i - channels] : 0;
                int estimate = left + up - corner, pa = Math.Abs(estimate - left), pb = Math.Abs(estimate - up), pc = Math.Abs(estimate - corner);
                var paeth = pa <= pb && pa <= pc ? left : pb <= pc ? up : corner;
                var predicted = filter == 1 ? left : filter == 2 ? up : filter == 3 ? (left + up) / 2 : filter == 4 ? paeth : 0;
                raw.WriteByte((byte)(rows[y][i] - predicted));
            }
        }

        var data = raw.ToArray();
        using (var png = new MemoryStream())
        {
            png.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 8);
            var ihdr = new byte[13];
            Put(ihdr, 0, width);
            Put(ihdr, 4, height);
            ihdr[8] = 8;
            ihdr[9] = (byte)colorType;
            Chunk(png, "IHDR", ihdr);
            if (palette != null) { Chunk(png, "PLTE", palette); }
            if (transparency != null) { Chunk(png, "tRNS", transparency); }
            Chunk(png, "IDAT", Divit.Pdf.PdfFile.Deflate(data, data.Length));
            Chunk(png, "IEND", new byte[0]);
            return png.ToArray();
        }
    }

    /// <summary>
    /// The decompressed content stream of a page as text.
    /// </summary>
    internal static string Content(Divit.Pdf.PdfPage page) => Encoding.ASCII.GetString(page.Content.Data, 0, page.Content.Length);

    private static void Chunk(Stream png, string type, byte[] data)
    {
        var length = new byte[4];
        Put(length, 0, data.Length);
        png.Write(length, 0, 4);
        var body = Encoding.ASCII.GetBytes(type).Concat(data).ToArray();
        png.Write(body, 0, body.Length);
        var crc = new byte[4];
        Put(crc, 0, (int)Crc32(body));
        png.Write(crc, 0, 4);
    }

    private static void Put(byte[] buffer, int offset, int value)
    {
        buffer[offset] = (byte)(value >> 24);
        buffer[offset + 1] = (byte)(value >> 16);
        buffer[offset + 2] = (byte)(value >> 8);
        buffer[offset + 3] = (byte)value;
    }

    private static uint Crc32(byte[] data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in data)
        {
            crc ^= b;
            for (var k = 0; k < 8; k++)
            {
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
            }
        }

        return ~crc;
    }
}
