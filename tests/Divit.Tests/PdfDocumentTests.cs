namespace Divit.Tests;

using Divit.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using Reader = UglyToad.PdfPig.PdfDocument;

/// <summary>
/// The PDF writer's foundation: pages, embedded subset TrueType fonts, text, lines and shapes, images, file
/// properties, and reproducible files.
/// </summary>
public class PdfDocumentTests
{
    private const string Turkish = "Ölçüm raporu: Ayşe Yılmaz, ĞÜŞİÖÇ ğüşıöç";

    private static string Save(PdfDocument doc, string name, bool keep = false)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        doc.Save(path);
        Assert.Empty(PdfTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    private static PdfDocument NewDocument() => new PdfDocument { Created = XlsxTestKit.FixedTime, Title = "Ölçüm", Author = "Divit tests" };

    [Fact]
    public void TextComesBackOutWithItsFontAndPlace()
    {
        var doc = NewDocument();
        var font = doc.AddFont("Trebuchet MS");
        var bold = doc.AddFont("Trebuchet MS", bold: true);
        var page = doc.AddPage();
        page.DrawText(Turkish, 50, 100, font, 12);
        page.DrawText("Right edge", 545, 130, bold, 10, "#1F4E79", PdfAlign.Right);
        page.DrawText("Centred", 297.64, 160, font, 10, align: PdfAlign.Center);
        var path = Save(doc, "text.pdf");

        using (var reader = Reader.Open(path))
        {
            var read = reader.GetPage(1);
            Assert.Contains(Turkish, read.Text);
            var letters = read.Letters.ToList();
            var first = letters.First();
            Assert.Equal(50, first.StartBaseLine.X, 1);
            Assert.Equal(841.89 - 100, first.StartBaseLine.Y, 1);
            Assert.Matches("^[A-Z]{6}\\+TrebuchetMS$", first.FontName);

            var right = letters.Where(l => l.FontName?.EndsWith("TrebuchetMS-Bold", StringComparison.Ordinal) == true).ToList();
            Assert.Equal("Right edge", string.Concat(right.Select(l => l.Value)));
            Assert.Equal(545, right.Last().EndBaseLine.X, 1);
            Assert.Equal(font.Measure(Turkish, 12), letters.Where(l => Math.Abs(l.StartBaseLine.Y - 741.89) < 0.1).Last().EndBaseLine.X - 50, 1);

            var centred = letters.Where(l => Math.Abs(l.StartBaseLine.Y - (841.89 - 160)) < 0.1).ToList();
            Assert.Equal(297.64, (centred.First().StartBaseLine.X + centred.Last().EndBaseLine.X) / 2, 1);
        }
    }

    [Fact]
    public void FontsAreSubsetToTheGlyphsUsed()
    {
        var doc = NewDocument();
        var font = doc.AddFont("Trebuchet MS");
        doc.AddPage().DrawText("Şş", 50, 100, font, 12);
        var path = Save(doc, "subset.pdf");
        Assert.True(new FileInfo(path).Length < 30000, "file of " + new FileInfo(path).Length + " bytes");

        // The subset keeps glyph ids and metrics; only the used glyphs (and the parts of composites) have outlines.
        var face = font.Face;
        var glyphs = new SortedSet<ushort> { face.Glyph('Ş'), face.Glyph('ş') };
        var subset = new TrueTypeFont(FontSubset.Build(face, glyphs, new Dictionary<ushort, int> { [face.Glyph('Ş')] = 'Ş', [face.Glyph('ş')] = 'ş' }));
        Assert.Equal("Trebuchet MS", subset.Family);
        Assert.Equal(face.Glyph('ş'), subset.Glyph('ş'));
        Assert.Equal(face.Glyph('Ş'), subset.Glyph('Ş'));
        Assert.Equal(0, subset.Glyph('Z'));
        Assert.Equal(face.GlyphCount, subset.GlyphCount);
        Assert.Equal(face.Advances, subset.Advances);
        Assert.True(subset.TryTable("loca", out var loca, out _));
        int Outline(TrueTypeFont f, int g) => (int)(f.U32(loca + (4 * (g + 1))) - f.U32(loca + (4 * g)));
        Assert.True(Outline(subset, face.Glyph('ş')) > 0);
        Assert.True(Outline(subset, 0) > 0);
        Assert.Equal(0, Outline(subset, face.Glyph('Z')));
        // Ş and ş are composites: S and s with the cedilla; all their parts come along.
        Assert.Contains(face.Glyph('S'), glyphs);
        Assert.Contains(face.Glyph('s'), glyphs);
        Assert.Equal(6, glyphs.Count);
        Assert.True(subset.LongLoca);

        // The widths the PDF gives for the glyphs, in thousandths of the size.
        var pdf = File.ReadAllText(path);
        var widths = System.Text.RegularExpressions.Regex.Match(pdf, @"/W \[(.*?)\] >>").Groups[1].Value;
        var expected = string.Join(" ", new[] { face.Glyph('Ş'), face.Glyph('ş') }.OrderBy(g => g).Select(g => Math.Round(face.Advances[g] * 1000.0 / face.UnitsPerEm).ToString(System.Globalization.CultureInfo.InvariantCulture)));
        var written = System.Text.RegularExpressions.Regex.Matches(widths, @"\[([0-9 ]+)\]").Cast<System.Text.RegularExpressions.Match>().SelectMany(m => m.Groups[1].Value.Trim().Split(' '));
        Assert.Equal(expected, string.Join(" ", written));
    }

    [Fact]
    public void FontsAreFoundByNameFileOrCollection()
    {
        var doc = NewDocument();
        Assert.Equal("Regular", doc.AddFont("trebuchet ms").Style);
        Assert.Equal("Bold", doc.AddFont("Trebuchet MS", bold: true).Style);
        Assert.Equal("Italic", doc.AddFont("Trebuchet MS", italic: true).Style);
        Assert.Same(doc.AddFont("Trebuchet MS"), doc.AddFont("Trebuchet MS"));
        var fromFile = doc.AddFont(File.ReadAllBytes(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "trebuc.ttf")));
        Assert.Equal("Trebuchet MS", fromFile.Family);
        var collection = File.ReadAllBytes(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "cambria.ttc"));
        var math = doc.AddFont(collection, 1);
        Assert.Equal("Cambria Math", math.Family);
        doc.AddPage().DrawText("x = ∑ yᵢ", 50, 100, math, 14);
        Save(doc, "collection.pdf");

        Assert.Throws<ArgumentException>(() => doc.AddFont("No Such Font Family"));
        Assert.Throws<ArgumentException>(() => doc.AddFont(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 }));

        // An OpenType font with PostScript outlines: a directory of one table, CFF.
        var cff = new byte[64];
        "OTTO"u8.ToArray().CopyTo(cff, 0);
        cff[5] = 1;
        "CFF "u8.ToArray().CopyTo(cff, 12);
        cff[27] = 32;
        cff[31] = 4;
        Assert.Throws<NotSupportedException>(() => doc.AddFont(cff));
    }

    [Fact]
    public void ShapesAreDrawnWithTheLeastStateChanges()
    {
        var doc = NewDocument();
        var page = doc.AddPage();
        page.FillRectangle(10, 20, 100, 50, "#1F4E79").FillRectangle(10, 80, 100, 50, "#1f4e79");
        page.DrawLine(10, 200, 300, 200, "#C00000", 1, PdfDash.Dashed).DrawLine(10, 210, 300, 210, "#C00000", 2, PdfDash.Dashed);
        page.DrawLine(10, 220, 300, 220, "#C00000", 2);
        page.DrawPolyline(new[] { (10.0, 300.0), (50.0, 250.0), (90.0, 300.0) }, "#2E75B6", 1.5);
        page.FillPolygon(new[] { (100.0, 300.0), (150.0, 250.0), (200.0, 300.0) }, "#9DC3E6");
        page.FillEllipse(300, 300, 20, 10, "#FF0000").StrokeEllipse(300, 300, 25, 15, "#000000");
        page.StrokeRectangle(400, 400, 50, 50, "#000000", 0.25);
        Save(doc, "shapes.pdf");

        var content = PdfTestKit.Content(page);
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(content, "0.122 0.306 0.475 rg"));
        Assert.Contains("10 771.89 100 50 re f\n", content);
        Assert.Contains("[4 3 ] 0 d", content);
        Assert.Contains("[8 6 ] 0 d", content);
        Assert.Contains("[] 0 d", content);
        Assert.Contains("1 j 1 J", content);
        Assert.Contains("h f\n", content);
        Assert.Equal(8, System.Text.RegularExpressions.Regex.Matches(content, " c ").Count);
        Assert.Contains("0.25 w", content);
    }

    [Fact]
    public void ImagesKeepTheirDataWhenTheyCan()
    {
        var rgb = PdfTestKit.Png(4, 3, 2, (x, y) => new[] { (byte)(x * 60), (byte)(y * 80), (byte)200 });
        var gray = PdfTestKit.Png(4, 3, 0, (x, y) => new[] { (byte)(x * 60) });
        var keyed = PdfTestKit.Png(4, 3, 2, (x, y) => new[] { (byte)x, (byte)0, (byte)0 }, transparency: new byte[] { 0, 1, 0, 0, 0, 0 });

        var plain = PdfImage.From(rgb);
        Assert.False(plain.Compress);
        Assert.Contains("/Predictor 15 /Colors 3 /BitsPerComponent 8 /Columns 4", plain.Entries);
        Assert.Null(plain.Mask);
        Assert.Contains("/DeviceGray", PdfImage.From(gray).Entries);
        Assert.Contains("/Mask [1 1 0 0 0 0 ]", PdfImage.From(keyed).Entries);

        var jpeg = PdfImage.From(PdfTestKit.Jpeg);
        Assert.Contains("/DCTDecode", jpeg.Entries);
        Assert.Same(PdfTestKit.Jpeg, jpeg.Data);
        Assert.Equal((24, 16), (jpeg.Width, jpeg.Height));
    }

    [Fact]
    public void TransparencyBecomesASoftMask()
    {
        var rgba = PdfTestKit.Png(3, 6, 6, (x, y) => new[] { (byte)((10 * x) + (37 * y)), (byte)(20 * y), (byte)(7 + (x * y)), (byte)((x * 100) + y) });
        var image = PdfImage.From(rgba);
        Assert.True(image.Compress);
        var color = new List<byte>();
        var alpha = new List<byte>();
        for (var y = 0; y < 6; y++)
        {
            for (var x = 0; x < 3; x++)
            {
                color.AddRange(new[] { (byte)((10 * x) + (37 * y)), (byte)(20 * y), (byte)(7 + (x * y)) });
                alpha.Add((byte)((x * 100) + y));
            }
        }

        Assert.Equal(color.ToArray(), image.Data);
        Assert.Equal(alpha.ToArray(), image.Mask!.Data);

        var grayAlpha = PdfImage.From(PdfTestKit.Png(2, 1, 4, (x, y) => new[] { (byte)(50 + x), (byte)(250 - x) }));
        Assert.Equal(new byte[] { 50, 51 }, grayAlpha.Data);
        Assert.Equal(new byte[] { 250, 249 }, grayAlpha.Mask!.Data);

        var palette = new byte[] { 255, 0, 0, 0, 255, 0, 0, 0, 255 };
        var indexed = PdfImage.From(PdfTestKit.Png(3, 1, 3, (x, y) => new[] { (byte)x }, palette, new byte[] { 0, 128 }));
        Assert.Contains("/Indexed /DeviceRGB 2 <FF000000FF000000FF>", indexed.Entries);
        Assert.Equal(new byte[] { 0, 128, 255 }, indexed.Mask!.Data);
    }

    [Fact]
    public void ImagesAreStoredOnceAndDrawnInBoxes()
    {
        var doc = NewDocument();
        var png = PdfTestKit.Png(40, 20, 6, (x, y) => new[] { (byte)200, (byte)(x * 6), (byte)50, (byte)(y * 12) });
        var first = doc.AddPage();
        first.DrawImage(png, 50, 60, 200).DrawImage(png, 50, 200, 100, 30).DrawImage(PdfTestKit.Jpeg, 300, 60, 120);
        doc.AddPage().DrawImage(png, 10, 10, 50);
        var path = Save(doc, "images.pdf");

        Assert.Single(System.Text.RegularExpressions.Regex.Matches(File.ReadAllText(path), "/SMask"));
        using (var reader = Reader.Open(path))
        {
            var images = reader.GetPage(1).GetImages().ToList();
            Assert.Equal(3, images.Count);
            var box = images[0].Bounds;
            Assert.Equal(50, box.Left, 1);
            Assert.Equal(200, box.Width, 1);
            Assert.Equal(100, box.Height, 1);
            Assert.Equal(841.89 - 60, box.Top, 1);
            Assert.Equal(30, images[1].Bounds.Height, 1);
            Assert.Single(reader.GetPage(2).GetImages());
        }

        Assert.Throws<NotSupportedException>(() => doc.AddPage().DrawImage(TestImages.Logo().Length > 0 ? Gif() : Gif(), 0, 0, 10));
        Assert.Throws<ArgumentException>(() => doc.AddPage().DrawImage(new byte[] { 1, 2, 3 }, 0, 0, 10));
    }

    [Fact]
    public void PagesHaveTheirSizes()
    {
        var doc = NewDocument();
        doc.AddPage();
        doc.AddPage(PdfPageSize.A3, landscape: true);
        doc.AddPage(PdfPageSize.Letter);
        doc.AddPage(200, 100);
        var path = Save(doc, "sizes.pdf");
        using (var reader = Reader.Open(path))
        {
            Assert.Equal(4, reader.NumberOfPages);
            Assert.Equal((595.28, 841.89), (reader.GetPage(1).Width, reader.GetPage(1).Height));
            Assert.Equal((1190.55, 841.89), (reader.GetPage(2).Width, reader.GetPage(2).Height));
            Assert.Equal((612.0, 792.0), (reader.GetPage(3).Width, reader.GetPage(3).Height));
            Assert.Equal((200.0, 100.0), (reader.GetPage(4).Width, reader.GetPage(4).Height));
        }
    }

    [Fact]
    public void PropertiesAreReadable()
    {
        var doc = NewDocument();
        doc.Subject = "Antenna";
        doc.Keywords = "SH2000; VSWR";
        doc.AddPage();
        var path = Save(doc, "properties.pdf");
        using (var reader = Reader.Open(path))
        {
            var info = reader.Information;
            Assert.Equal("Ölçüm", info.Title);
            Assert.Equal("Divit tests", info.Author);
            Assert.Equal("Antenna", info.Subject);
            Assert.Equal("SH2000; VSWR", info.Keywords);
            Assert.Equal("Divit", info.Creator);
            Assert.Equal("D:20260102030405Z", info.CreationDate);
        }
    }

    [Fact]
    public void NumbersAreWrittenWithoutExponents()
    {
        var b = new PdfBuffer();
        foreach (var value in new[] { 0, 1.5, -2.25, 1.23456, 0.0004, -0.0004, 1e6, 841.89, 0.1 + 0.2 })
        {
            b.Number(value);
        }

        Assert.Equal("0 1.5 -2.25 1.235 0 0 1000000 841.89 0.3 ", System.Text.Encoding.ASCII.GetString(b.Data, 0, b.Length));
    }

    [Fact]
    public void CompressedDataInflatesBack()
    {
        var data = System.Text.Encoding.ASCII.GetBytes(string.Concat(Enumerable.Repeat("BT /F1 12 Tf 50 741.89 Td <0041> Tj ET\n", 2000)));
        var packed = PdfFile.Deflate(data, data.Length);
        Assert.Equal(0x78, packed[0]);
        using (var inflate = new System.IO.Compression.DeflateStream(new MemoryStream(packed, 2, packed.Length - 6), System.IO.Compression.CompressionMode.Decompress))
        using (var back = new MemoryStream())
        {
            inflate.CopyTo(back);
            Assert.Equal(data, back.ToArray());
        }

        uint a = 1, s = 0;
        foreach (var d in data)
        {
            a = (a + d) % 65521;
            s = (s + a) % 65521;
        }

        var adler = (uint)((packed[packed.Length - 4] << 24) | (packed[packed.Length - 3] << 16) | (packed[packed.Length - 2] << 8) | packed[packed.Length - 1]);
        Assert.Equal((s << 16) | a, adler);
    }

    [Fact]
    public void MissingGlyphsAndEmptyPagesStayValid()
    {
        var doc = NewDocument();
        var font = doc.AddFont("Trebuchet MS");
        doc.AddPage().DrawText("\U0001F600 ☃", 50, 100, font, 12);
        doc.AddPage();
        Save(doc, "missing.pdf");
        Assert.False(font.Has('☃'));
        Assert.True(font.Has('ş'));
    }

    [Fact]
    public void PairsAreKerned()
    {
        var doc = NewDocument();
        var font = doc.AddFont("Trebuchet MS");
        var face = font.Face;
        Assert.True(font.Measure("AV", 100) < font.Measure("A", 100) + font.Measure("V", 100) - 1, "AV is kerned");
        Assert.True(font.Measure("To", 100) < font.Measure("T", 100) + font.Measure("o", 100) - 1, "To is kerned");
        Assert.Equal(font.Measure("A V", 100), font.Measure("A", 100) + font.Measure(" ", 100) + font.Measure("V", 100), 6);

        // The old kern table, read when a font has no GPOS kerning, knows the same pair.
        var old = FontKerning.Read(face, gpos: false);
        Assert.True(old.Any);
        Assert.True(old.Pair(face.Glyph('A'), face.Glyph('V')) < 0);
        Assert.True(face.Kerning.Pair(face.Glyph('A'), face.Glyph('V')) < 0);

        var page = doc.AddPage();
        page.DrawText("AVATAR Today", 50, 100, font, 40);
        Assert.Contains("] TJ", PdfTestKit.Content(page));
        var path = Save(doc, "kerning.pdf");
        using (var reader = Reader.Open(path))
        {
            var letters = reader.GetPage(1).Letters.ToList();
            var a = letters[0];
            var v = letters[1];
            Assert.Equal("A", a.Value);
            Assert.Equal(50 + font.Measure("AV", 40) - font.Measure("V", 40), v.StartBaseLine.X, 1);
            Assert.Equal(50 + font.Measure("AVATAR Today", 40), letters.Last().EndBaseLine.X, 1);
        }
    }

    [Fact]
    public void DocumentsAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        foreach (var name in new[] { "a.pdf", "b.pdf" })
        {
            var doc = NewDocument();
            var page = doc.AddPage();
            page.DrawText(Turkish, 50, 100, doc.AddFont("Trebuchet MS"), 12).DrawImage(PdfTestKit.Jpeg, 50, 120, 100);
            doc.Save(Path.Combine(folder, name));
        }

        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.pdf")), File.ReadAllBytes(Path.Combine(folder, "b.pdf")));
    }

    private static byte[] Gif() => new byte[] { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 1, 0, 1, 0, 0, 0, 0 };
}
