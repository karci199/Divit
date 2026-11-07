namespace Divit.Tests;

using Divit.Packaging;
using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Xunit;

/// <summary>
/// The package layer every format builds on: zip parts, content types, relationships, properties, the atomic save
/// and the allocation-free number formatting.
/// </summary>
public class PackagingTests
{
    private static ZipArchive Package(Action<OpcPackage> write, DateTime? time = null)
    {
        var stream = new MemoryStream();
        OpcPackage.Write(stream, time ?? XlsxTestKit.FixedTime, write);
        stream.Position = 0;
        return new ZipArchive(stream, ZipArchiveMode.Read);
    }

    private static XDocument Xml(ZipArchive zip, string name)
    {
        using (var stream = zip.GetEntry(name)!.Open())
        {
            return XDocument.Load(stream);
        }
    }

    [Fact]
    public void PartsComeInOrderAndContentTypesLast()
    {
        using (var zip = Package(p =>
        {
            p.XmlPart("a/one.xml", w => w.WriteElementString("one", "1"));
            p.BinaryPart("a/two.bin", new byte[] { 1, 2, 3 }, true);
            p.Default("xml", ContentTypes.Xml);
            p.Default("bin", "application/octet-stream");
            p.Override("/a/one.xml", "application/x-one+xml");
        }))
        {
            Assert.Equal(new[] { "a/one.xml", "a/two.bin", "[Content_Types].xml" }, zip.Entries.Select(e => e.FullName).ToArray());

            var types = Xml(zip, "[Content_Types].xml").Root!;
            Assert.Equal("http://schemas.openxmlformats.org/package/2006/content-types", types.Name.NamespaceName);
            Assert.Equal(new[] { "xml", "bin" }, types.Elements().Where(e => e.Name.LocalName == "Default").Select(e => (string)e.Attribute("Extension")!).ToArray());
            Assert.Equal("/a/one.xml", (string)types.Elements().Single(e => e.Name.LocalName == "Override").Attribute("PartName")!);
        }
    }

    [Fact]
    public void XmlPartsAreUtf8WithoutByteOrderMark()
    {
        using (var zip = Package(p =>
        {
            p.XmlPart("with.xml", w => w.WriteElementString("t", "Ölçüm"));
            p.XmlPart("without.vml", w => w.WriteElementString("t", "x"), declaration: false);
        }))
        {
            byte[] Bytes(string name)
            {
                using (var stream = zip.GetEntry(name)!.Open())
                using (var copy = new MemoryStream())
                {
                    stream.CopyTo(copy);
                    return copy.ToArray();
                }
            }

            var with = Bytes("with.xml");
            Assert.NotEqual(0xEF, with[0]);
            Assert.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>", System.Text.Encoding.UTF8.GetString(with));
            Assert.Contains("Ölçüm", System.Text.Encoding.UTF8.GetString(with));
            Assert.Equal("<t>x</t>", System.Text.Encoding.UTF8.GetString(Bytes("without.vml")));
        }
    }

    [Fact]
    public void EveryEntryCarriesTheGivenTime()
    {
        using (var zip = Package(p =>
        {
            p.XmlPart("a.xml", w => w.WriteElementString("a", "1"));
            p.BinaryPart("b.png", new byte[] { 1 }, false);
        }, new DateTime(2030, 6, 15, 10, 20, 30, DateTimeKind.Utc)))
        {
            Assert.All(zip.Entries, e => Assert.Equal(new DateTime(2030, 6, 15, 10, 20, 30), e.LastWriteTime.DateTime));
        }
    }

    [Fact]
    public void TimesOutsideTheZipRangeFallBackToTheClock()
    {
        using (var zip = Package(p => p.XmlPart("a.xml", w => w.WriteElementString("a", "1")), new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)))
        {
            Assert.InRange(zip.Entries[0].LastWriteTime.Year, 2020, 2107);
        }
    }

    [Fact]
    public void EmptyRelationshipListsWriteNoPart()
    {
        using (var zip = Package(p =>
        {
            p.Relationships("_rels/empty.rels", new RelationshipList());
            var list = new RelationshipList();
            list.Add(RelationshipTypes.Image, "media/a.png");
            p.Relationships("_rels/full.rels", list);
        }))
        {
            Assert.Null(zip.GetEntry("_rels/empty.rels"));
            Assert.NotNull(zip.GetEntry("_rels/full.rels"));
        }
    }

    [Fact]
    public void RelationshipsNumberThemselvesAndMarkExternalTargets()
    {
        var list = new RelationshipList();
        Assert.Equal("rId1", list.Add(RelationshipTypes.Worksheet, "worksheets/sheet1.xml"));
        Assert.Equal("rId2", list.Add(RelationshipTypes.Hyperlink, "https://example.com/?a=1&b=2", true));
        Assert.Equal(2, list.Count);

        using (var zip = Package(p => p.Relationships("_rels/.rels", list)))
        {
            var items = Xml(zip, "_rels/.rels").Root!.Elements().ToList();
            Assert.Null(items[0].Attribute("TargetMode"));
            Assert.Equal("External", (string)items[1].Attribute("TargetMode")!);
            Assert.Equal("https://example.com/?a=1&b=2", (string)items[1].Attribute("Target")!);
            Assert.Equal(RelationshipTypes.Hyperlink, (string)items[1].Attribute("Type")!);
        }
    }

    [Fact]
    public void DocumentPropertiesFallBackToTheApplicationAndDropControlCharacters()
    {
        using (var zip = Package(p => p.DocumentProperties(new DocumentProperties("Divit")
        {
            Title = "Report\u0001 2026",
            Created = new DateTime(2026, 3, 4, 5, 6, 7, DateTimeKind.Utc),
        })))
        {
            var core = Xml(zip, "docProps/core.xml").Root!;
            XNamespace dc = "http://purl.org/dc/elements/1.1/";
            XNamespace terms = "http://purl.org/dc/terms/";
            Assert.Equal("Report 2026", (string)core.Element(dc + "title")!);
            Assert.Equal("Divit", (string)core.Element(dc + "creator")!);
            Assert.Null(core.Element(dc + "subject"));
            Assert.Equal("2026-03-04T05:06:07Z", (string)core.Element(terms + "created")!);
            Assert.Equal("2026-03-04T05:06:07Z", (string)core.Element(terms + "modified")!);
            Assert.Equal("Divit", Xml(zip, "docProps/app.xml").Root!.Elements().Single().Value);
        }
    }

    [Fact]
    public void AFailedSaveKeepsTheOldFileAndLeavesNoTemporaryFile()
    {
        var folder = XlsxTestKit.NewFolder();
        var path = Path.Combine(folder, "doc.zip");
        File.WriteAllText(path, "previous version");

        Assert.Throws<InvalidOperationException>(() => OpcPackage.Save(path, null, p =>
        {
            p.XmlPart("a.xml", w => w.WriteElementString("a", "1"));
            throw new InvalidOperationException("stop");
        }));

        Assert.Equal("previous version", File.ReadAllText(path));
        Assert.Equal(new[] { "doc.zip" }, Directory.GetFiles(folder).Select(Path.GetFileName).ToArray());
    }

    [Fact]
    public void ASaveReplacesTheOldFile()
    {
        var folder = XlsxTestKit.NewFolder();
        var path = Path.Combine(folder, "doc.zip");
        File.WriteAllText(path, "previous version");
        OpcPackage.Save(path, XlsxTestKit.FixedTime, p => p.XmlPart("a.xml", w => w.WriteElementString("a", "1")));

        using (var zip = ZipFileOpen(path))
        {
            Assert.NotNull(zip.GetEntry("a.xml"));
        }

        Assert.Single(Directory.GetFiles(folder));
    }

    private static ZipArchive ZipFileOpen(string path) => new ZipArchive(File.OpenRead(path), ZipArchiveMode.Read);

    [Theory]
    [InlineData("#1f4e9a", "1F4E9A")]
    [InlineData("1F4E9A", "1F4E9A")]
    [InlineData("#000000", "000000")]
    public void ColoursAreNormalised(string color, string expected) => Assert.Equal(expected, OfficeText.Rgb(color));

    [Theory]
    [InlineData("")]
    [InlineData("#12345")]
    [InlineData("#1234567")]
    [InlineData("#GG0000")]
    [InlineData("red")]
    public void BadColoursAreRejected(string color) => Assert.Throws<ArgumentException>(() => OfficeText.Rgb(color));

    [Fact]
    public void PngSignatureIsChecked()
    {
        Assert.True(OfficeText.IsPng(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0 }));
        Assert.False(OfficeText.IsPng(new byte[] { 0x89, 0x50, 0x4E, 0x47 }));
        Assert.False(OfficeText.IsPng(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0, 0x10, 0x4A, 0x46 }));
    }

    [Fact]
    public void CleanTextIsReturnedAsItIs()
    {
        var text = "tab\tline\nreturn\r";
        Assert.Same(text, OfficeText.Clean(text));
        Assert.Equal("ab", OfficeText.Clean("\u0001a\u0000b\u001F"));
        Assert.Equal(string.Empty, OfficeText.Clean("\u0002"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(7)]
    [InlineData(10)]
    [InlineData(99)]
    [InlineData(1048576)]
    [InlineData(int.MaxValue)]
    public void WholeNumbersFormatLikeToString(int value)
    {
        var buffer = new char[32];
        buffer[0] = 'x';
        var end = OfficeText.FormatInt(buffer, 1, value);
        Assert.Equal("x" + value.ToString(CultureInfo.InvariantCulture), new string(buffer, 0, end));
    }

    [Fact]
    public void NumbersFormatExactlyLikeRoundTrip()
    {
        var values = new[]
        {
            0.0, -0.0, 1, -1, 0.1, -0.1, 1.5, 70.3, 1e14, 999999999999999, 1e15, -1e15, 1e16, 1.0 / 3, 2.0 / 3, Math.PI,
            -Math.E, 1e-5, 1e-300, 1e300, double.MaxValue, double.MinValue, double.Epsilon, 123456789.123456789,
            4.35, 0.30000000000000004, long.MaxValue, (double)int.MinValue,
        };
        var random = new Random(42);
        var all = values.Concat(Enumerable.Range(0, 2000).Select(_ => (random.NextDouble() - 0.5) * Math.Pow(10, random.Next(-20, 20))))
            .Concat(Enumerable.Range(0, 500).Select(i => (double)random.Next(-1000000, 1000000)));
        var buffer = new char[32];
        foreach (var value in all)
        {
            var length = OfficeText.FormatDouble(buffer, value);
            Assert.Equal(value.ToString("R", CultureInfo.InvariantCulture), new string(buffer, 0, length));
        }
    }
}
