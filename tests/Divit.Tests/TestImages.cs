namespace Divit.Tests;

using System;
using System.IO;
using System.IO.Compression;

/// <summary>
/// Pictures for the demo documents: drawn with System.Drawing on .NET Framework (text, shapes), and as simple
/// generated PNGs elsewhere, so every target can save the documents.
/// </summary>
internal static class TestImages
{
    /// <summary>
    /// A PNG of the given size whose pixels come from a function returning 0xRRGGBB.
    /// </summary>
    internal static byte[] Png(int width, int height, Func<int, int, int> pixel)
    {
        var raw = new byte[(width * 3 + 1) * height];
        var i = 0;
        for (var y = 0; y < height; y++)
        {
            raw[i++] = 0;
            for (var x = 0; x < width; x++)
            {
                var c = pixel(x, y);
                raw[i++] = (byte)(c >> 16);
                raw[i++] = (byte)(c >> 8);
                raw[i++] = (byte)c;
            }
        }

        using (var png = new MemoryStream())
        {
            png.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 8);
            var header = new byte[13];
            BigEndian(header, 0, width);
            BigEndian(header, 4, height);
            header[8] = 8;
            header[9] = 2;
            Chunk(png, "IHDR", header);
            using (var deflated = new MemoryStream())
            {
                deflated.WriteByte(0x78);
                deflated.WriteByte(0x01);
                using (var deflate = new DeflateStream(deflated, CompressionLevel.Fastest, true))
                {
                    deflate.Write(raw, 0, raw.Length);
                }

                var adler = Adler32(raw);
                var tail = new byte[4];
                BigEndian(tail, 0, (int)adler);
                deflated.Write(tail, 0, 4);
                Chunk(png, "IDAT", deflated.ToArray());
            }

            Chunk(png, "IEND", new byte[0]);
            return png.ToArray();
        }
    }

    /// <summary>
    /// A logo: a blue square with a white wave, in the house colours.
    /// </summary>
    internal static byte[] Logo() =>
#if NETFRAMEWORK
        Draw(360, 120, g =>
        {
            using (var blue = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(31, 78, 121)))
            using (var light = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(157, 195, 230)))
            using (var font = new System.Drawing.Font("Trebuchet MS", 34, System.Drawing.FontStyle.Bold))
            using (var small = new System.Drawing.Font("Trebuchet MS", 12))
            using (var pen = new System.Drawing.Pen(System.Drawing.Color.White, 6))
            {
                g.FillRectangle(blue, 6, 10, 100, 100);
                g.DrawArc(pen, 26, 30, 60, 60, 200, 140);
                g.DrawArc(pen, 38, 42, 36, 36, 200, 140);
                g.FillEllipse(System.Drawing.Brushes.White, 50, 72, 12, 12);
                g.DrawString("DIVIT", font, blue, 116, 14);
                g.DrawString("ELECTRONICS", small, light, 122, 72);
            }
        });
#else
        Png(360, 120, (x, y) => x < 110 ? 0x1F4E79 : (y > 50 && y < 70 ? 0x9DC3E6 : 0xFFFFFF));
#endif

    /// <summary>
    /// The unit under test: a horn antenna seen from the side.
    /// </summary>
    internal static byte[] Antenna() =>
#if NETFRAMEWORK
        Draw(520, 300, g =>
        {
            using (var metal = new System.Drawing.Drawing2D.LinearGradientBrush(new System.Drawing.Rectangle(0, 60, 520, 180), System.Drawing.Color.FromArgb(220, 224, 230), System.Drawing.Color.FromArgb(120, 130, 145), 90f))
            using (var outline = new System.Drawing.Pen(System.Drawing.Color.FromArgb(60, 70, 85), 3))
            using (var font = new System.Drawing.Font("Trebuchet MS", 14))
            {
                var horn = new[] { new System.Drawing.Point(120, 130), new System.Drawing.Point(420, 40), new System.Drawing.Point(420, 260), new System.Drawing.Point(120, 170) };
                g.FillPolygon(metal, horn);
                g.DrawPolygon(outline, horn);
                g.FillRectangle(metal, 60, 125, 60, 50);
                g.DrawRectangle(outline, 60, 125, 60, 50);
                g.FillRectangle(System.Drawing.Brushes.Goldenrod, 30, 138, 30, 24);
                g.DrawString("SH2000 dual-ridge horn", font, System.Drawing.Brushes.DimGray, 230, 268);
                g.DrawString("SMA connector", font, System.Drawing.Brushes.DimGray, 4, 186);
            }
        });
#else
        Png(520, 300, (x, y) => x > 120 && x < 420 && Math.Abs(y - 150) < 20 + (x - 120) / 3 ? 0x8C96A0 : 0xFFFFFF);
#endif

    /// <summary>
    /// The test set-up: network analyser, cable, antenna, absorber.
    /// </summary>
    internal static byte[] SetUp() =>
#if NETFRAMEWORK
        Draw(700, 300, g =>
        {
            using (var box = new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(222, 235, 247)))
            using (var line = new System.Drawing.Pen(System.Drawing.Color.FromArgb(31, 78, 121), 3))
            using (var cable = new System.Drawing.Pen(System.Drawing.Color.FromArgb(197, 90, 17), 4))
            using (var font = new System.Drawing.Font("Trebuchet MS", 13))
            using (var bold = new System.Drawing.Font("Trebuchet MS", 13, System.Drawing.FontStyle.Bold))
            {
                var ink = System.Drawing.Brushes.Black;
                g.FillRectangle(box, 20, 90, 180, 120);
                g.DrawRectangle(line, 20, 90, 180, 120);
                g.DrawString("Vector network", bold, ink, 40, 115);
                g.DrawString("analyser", bold, ink, 70, 140);
                g.DrawString("Port 1", font, ink, 140, 182);
                g.DrawBezier(cable, 200, 190, 280, 260, 340, 120, 420, 150);
                g.DrawString("Calibrated cable", font, ink, 250, 230);
                g.FillPolygon(System.Drawing.Brushes.LightSteelBlue, new[] { new System.Drawing.Point(420, 140), new System.Drawing.Point(540, 90), new System.Drawing.Point(540, 210), new System.Drawing.Point(420, 160) });
                g.DrawPolygon(line, new[] { new System.Drawing.Point(420, 140), new System.Drawing.Point(540, 90), new System.Drawing.Point(540, 210), new System.Drawing.Point(420, 160) });
                g.DrawString("SH2000", bold, ink, 445, 215);
                for (var i = 0; i < 6; i++)
                {
                    g.FillPolygon(System.Drawing.Brushes.DimGray, new[] { new System.Drawing.Point(640, 60 + (i * 35)), new System.Drawing.Point(690, 77 + (i * 35)), new System.Drawing.Point(640, 95 + (i * 35)) });
                }

                g.DrawString("Absorber", font, ink, 600, 275);
            }
        });
#else
        Png(700, 300, (x, y) => (x > 20 && x < 200 && y > 90 && y < 210) || (x > 420 && x < 540 && y > 100 && y < 200) ? 0xDEEBF7 : 0xFFFFFF);
#endif

#if NETFRAMEWORK
    private static byte[] Draw(int width, int height, Action<System.Drawing.Graphics> draw)
    {
        using (var bitmap = new System.Drawing.Bitmap(width, height))
        using (var g = System.Drawing.Graphics.FromImage(bitmap))
        using (var stream = new MemoryStream())
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            g.Clear(System.Drawing.Color.White);
            draw(g);
            bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
            return stream.ToArray();
        }
    }
#endif

    private static void Chunk(Stream png, string type, byte[] data)
    {
        var length = new byte[4];
        BigEndian(length, 0, data.Length);
        png.Write(length, 0, 4);
        var body = new byte[4 + data.Length];
        for (var i = 0; i < 4; i++)
        {
            body[i] = (byte)type[i];
        }

        Array.Copy(data, 0, body, 4, data.Length);
        png.Write(body, 0, body.Length);
        var crc = new byte[4];
        BigEndian(crc, 0, (int)Crc32(body));
        png.Write(crc, 0, 4);
    }

    private static void BigEndian(byte[] buffer, int offset, int value)
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

    private static uint Adler32(byte[] data)
    {
        uint a = 1, b = 0;
        foreach (var d in data)
        {
            a = (a + d) % 65521;
            b = (b + a) % 65521;
        }

        return (b << 16) | a;
    }
}
