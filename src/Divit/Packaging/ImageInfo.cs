#nullable enable
namespace Divit.Packaging;

/// <summary>
/// The format and pixel size of an image, read from its first bytes: PNG, JPEG or GIF, the formats every Office
/// program shows. Shared by the formats that embed images.
/// </summary>
internal readonly struct ImageInfo
{
    private ImageInfo(string extension, string contentType, int width, int height)
    {
        Extension = extension;
        ContentType = contentType;
        Width = width;
        Height = height;
    }

    /// <summary>
    /// File extension of the media part: <c>png</c>, <c>jpeg</c> or <c>gif</c>.
    /// </summary>
    public string Extension { get; }

    /// <summary>
    /// Content type of the media part.
    /// </summary>
    public string ContentType { get; }

    /// <summary>
    /// Width in pixels; 0 when the header does not tell.
    /// </summary>
    public int Width { get; }

    /// <summary>
    /// Height in pixels; 0 when the header does not tell.
    /// </summary>
    public int Height { get; }

    /// <summary>
    /// Reads the format and size; <c>null</c> when the bytes are none of the three formats.
    /// </summary>
    public static ImageInfo? Read(byte[] data)
    {
        if (OfficeText.IsPng(data))
        {
            // The IHDR chunk follows the signature: width and height as 32-bit big-endian numbers.
            return new ImageInfo("png", "image/png", data.Length >= 24 ? BigEndian(data, 16, 4) : 0, data.Length >= 24 ? BigEndian(data, 20, 4) : 0);
        }

        if (data.Length >= 10 && data[0] == 'G' && data[1] == 'I' && data[2] == 'F' && data[3] == '8')
        {
            return new ImageInfo("gif", "image/gif", data[6] | (data[7] << 8), data[8] | (data[9] << 8));
        }

        if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xD8)
        {
            var (width, height) = JpegSize(data);
            return new ImageInfo("jpeg", "image/jpeg", width, height);
        }

        return null;
    }

    /// <summary>
    /// Walks the JPEG markers to the first start-of-frame, which holds height and width.
    /// </summary>
    private static (int Width, int Height) JpegSize(byte[] data)
    {
        var i = 2;
        while (i + 9 < data.Length)
        {
            if (data[i] != 0xFF)
            {
                return (0, 0);
            }

            var marker = data[i + 1];
            if (marker == 0xFF)
            {
                // Fill byte before a marker.
                i++;
                continue;
            }

            // SOF0–SOF15, except DHT (C4), JPG (C8) and DAC (CC), which share the range.
            if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
            {
                return (BigEndian(data, i + 7, 2), BigEndian(data, i + 5, 2));
            }

            if (marker == 0xD8 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01)
            {
                // Markers without a length.
                i += 2;
                continue;
            }

            i += 2 + BigEndian(data, i + 2, 2);
        }

        return (0, 0);
    }

    private static int BigEndian(byte[] data, int offset, int count)
    {
        var value = 0;
        for (var k = 0; k < count; k++)
        {
            value = (value << 8) | data[offset + k];
        }

        return value;
    }
}
