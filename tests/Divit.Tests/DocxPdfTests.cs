namespace Divit.Tests;

using Divit.Docx;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using UglyToad.PdfPig.Content;
using Xunit;
using Reader = UglyToad.PdfPig.PdfDocument;

/// <summary>
/// Word documents laid out as PDF: the demos saved as PDF and kept for a look, and the layout rules checked on the
/// positions of the words read back.
/// </summary>
public class DocxPdfTests
{
    /// <summary>
    /// The right margin of an A4 page with 2.5 cm margins, in points.
    /// </summary>
    private const double RightMargin = 595.28 - (2.5 * 72 / 2.54);

    private const double LeftMargin = 2.5 * 72 / 2.54;

    private static string Save(DocxDocument doc, string name, bool keep = false)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        doc.SavePdf(path);
        Assert.Empty(PdfTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    private static DocxDocument NewDocument() => new DocxDocument { Created = XlsxTestKit.FixedTime, NumberHeadings = true, Title = "Layout" };

    private static List<List<Word>> Words(string path)
    {
        using (var reader = Reader.Open(path))
        {
            return reader.GetPages().Select(p => p.GetWords().ToList()).ToList();
        }
    }

    [Fact]
    public void TheDemosAreLaidOut()
    {
        var path = Save(DocxProcedureTests.Procedure(), "acceptance-test-procedure.pdf", keep: true);
        using (var reader = Reader.Open(path))
        {
            Assert.Equal(10, reader.NumberOfPages);
            Assert.Contains("SH2000", reader.GetPage(1).Text);
        }

        Save(DocxFeatureTests.Showcase(), "word-features.pdf", keep: true);
        Save(DocxTechnicalTests.Showcase(), "word-equations-sources.pdf", keep: true);
    }

    [Fact]
    public void ContentsShowTheRealPagesAtTheRightMargin()
    {
        var doc = NewDocument();
        doc.AddTableOfContents(2, "Contents");
        foreach (var name in new[] { "Alpha", "Beta", "Gamma" })
        {
            doc.AddPageBreak();
            doc.AddHeading(name);
            doc.AddParagraph("Text of " + name + ".");
        }

        var pages = Words(Save(doc, "contents.pdf"));
        Assert.Equal(4, pages.Count);
        for (var i = 0; i < 3; i++)
        {
            var name = new[] { "Alpha", "Beta", "Gamma" }[i];
            var entry = pages[0].Single(w => w.Text == name);
            var page = pages[0].Where(w => Math.Abs(w.BoundingBox.Bottom - entry.BoundingBox.Bottom) < 1).OrderBy(w => w.BoundingBox.Right).Last();
            Assert.Equal((i + 2).ToString(CultureInfo.InvariantCulture), page.Text);
            Assert.Equal(RightMargin, page.BoundingBox.Right, 0);
            Assert.Contains(pages[i + 1], w => w.Text == name);
        }
    }

    [Fact]
    public void HeaderRowsRepeatAndNoRowIsLostOrDoubled()
    {
        var doc = NewDocument();
        var table = doc.AddTable(2);
        table.AddHeaderRow("Point", "Value");
        for (var i = 1; i <= 120; i++)
        {
            table.AddRow("R" + i.ToString(CultureInfo.InvariantCulture), "V");
        }

        var pages = Words(Save(doc, "long-table.pdf"));
        Assert.True(pages.Count >= 3, pages.Count + " pages");
        foreach (var page in pages)
        {
            var header = page.Single(w => w.Text == "Point");
            Assert.True(page.Where(w => w.Text.StartsWith("R", StringComparison.Ordinal)).All(w => w.BoundingBox.Top < header.BoundingBox.Bottom), "header row on top");
        }

        var rows = pages.SelectMany(p => p).Where(w => w.Text.StartsWith("R", StringComparison.Ordinal)).Select(w => w.Text).ToList();
        Assert.Equal(Enumerable.Range(1, 120).Select(i => "R" + i.ToString(CultureInfo.InvariantCulture)), rows);
    }

    [Fact]
    public void ListsAndHeadingsHangTheirMarkers()
    {
        var doc = NewDocument();
        doc.AddHeading("Scope");
        var steps = doc.AddList(DocxListKind.Numbered);
        steps.AddItem("First");
        steps.AddItem("Inner", 1);
        steps.AddItem("Second");
        steps.AddItem("Again", 1);
        doc.AddList().AddItem("Bullet");
        var words = Words(Save(doc, "lists.pdf"))[0];

        Word At(string text) => words.First(w => w.Text == text);
        Assert.Equal(LeftMargin, At("1").BoundingBox.Left, 0);
        Assert.Equal(LeftMargin + 21.6, At("Scope").BoundingBox.Left, 0);
        Assert.Equal(new[] { "1.", "a)", "2.", "a)", "•" }, words.Where(w => w.Text is "1." or "a)" or "b)" or "2." or "•").Select(w => w.Text));
        Assert.Equal(LeftMargin + 36, At("First").BoundingBox.Left, 0);
        Assert.Equal(LeftMargin + 72, At("Inner").BoundingBox.Left, 0);
        Assert.Equal(At("First").BoundingBox.Left, At("Second").BoundingBox.Left, 1);
        Assert.Equal(LeftMargin + 18, At("1.").BoundingBox.Left, 0);
    }

    [Fact]
    public void JustifiedLinesFillTheWidth()
    {
        var doc = NewDocument();
        var paragraph = doc.AddParagraph(string.Join(" ", Enumerable.Repeat("measured gain of the horn antenna", 12)));
        paragraph.Alignment = DocxAlignment.Justify;
        var words = Words(Save(doc, "justify.pdf"))[0];
        var lines = words.GroupBy(w => Math.Round(w.BoundingBox.Bottom)).OrderByDescending(g => g.Key).ToList();
        Assert.True(lines.Count >= 3);
        foreach (var line in lines.Take(lines.Count - 1))
        {
            Assert.Equal(LeftMargin, line.Min(w => w.BoundingBox.Left), 0);
            Assert.Equal(RightMargin, line.Max(w => w.BoundingBox.Right), 0);
        }

        Assert.True(lines.Last().Max(w => w.BoundingBox.Right) < RightMargin - 20, "the last line is not stretched");
    }

    [Fact]
    public void PageFieldsCountThePages()
    {
        var doc = NewDocument();
        doc.Section.Header.AddParagraph().Add("Page ").AddField(DocxField.Page).Add(" of ").AddField(DocxField.NumPages);
        doc.Section.Watermark = "DRAFT";
        for (var i = 0; i < 3; i++)
        {
            doc.AddParagraph("Body " + i.ToString(CultureInfo.InvariantCulture));
            doc.AddPageBreak();
        }

        var path = Save(doc, "fields.pdf");
        using (var reader = Reader.Open(path))
        {
            Assert.Equal(4, reader.NumberOfPages);
            for (var i = 1; i <= 4; i++)
            {
                var text = reader.GetPage(i).Text;
                Assert.Contains("Page " + i.ToString(CultureInfo.InvariantCulture) + " of 4", text);
                Assert.Contains("DRAFT", text);
            }
        }
    }

    [Fact]
    public void HeadingsStayWithTheirParagraph()
    {
        // Whatever space is left at the foot of the page, a heading never ends a page alone.
        for (var lines = 30; lines <= 50; lines++)
        {
            var doc = NewDocument();
            for (var i = 0; i < lines; i++)
            {
                doc.AddParagraph("Filler " + i.ToString(CultureInfo.InvariantCulture));
            }

            doc.AddHeading("Marker");
            doc.AddParagraph("Follower");
            var pages = Words(Save(doc, "keep.pdf"));
            var heading = pages.FindIndex(p => p.Any(w => w.Text == "Marker"));
            var follower = pages.FindIndex(p => p.Any(w => w.Text == "Follower"));
            Assert.Equal(heading, follower);
        }
    }

    [Fact]
    public void FootnotesSitAtTheFootOfTheirPage()
    {
        var doc = NewDocument();
        doc.AddParagraph().Add("Measured in the chamber.").AddFootnote("Anechoic chamber 2.");
        doc.AddParagraph("Next paragraph.");
        var words = Words(Save(doc, "notes.pdf"))[0];
        var note = words.Single(w => w.Text == "Anechoic");
        var body = words.Single(w => w.Text == "Next");
        Assert.True(note.BoundingBox.Bottom < 100, "note at the foot: " + note.BoundingBox.Bottom);
        Assert.True(note.BoundingBox.Bottom > (2.5 * 72 / 2.54) - 1, "note above the bottom margin");
        Assert.True(body.BoundingBox.Bottom > 700, "body at the top");
    }

    [Fact]
    public void LinksAndBookmarksAreThere()
    {
        var doc = NewDocument();
        doc.AddTableOfContents(1);
        doc.AddHeading("First");
        doc.AddParagraph().AddLink("site", "https://divit.example");
        doc.AddHeading("Second");
        var path = Save(doc, "links.pdf");
        using (var reader = Reader.Open(path))
        {
            var links = reader.GetPage(1).GetHyperlinks().ToList();
            Assert.Contains(links, l => l.Uri == "https://divit.example");
            Assert.True(reader.TryGetBookmarks(out var bookmarks));
            Assert.Equal(new[] { "1 First", "2 Second" }, bookmarks.Roots.Select(b => b.Title));
        }

        var text = File.ReadAllText(path);
        Assert.Equal(3, System.Text.RegularExpressions.Regex.Matches(text, "/Subtype /Link").Count);
    }

    [Fact]
    public void PdfsOfWordDocumentsAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        DocxFeatureTests.Showcase().SavePdf(Path.Combine(folder, "a.pdf"));
        DocxFeatureTests.Showcase().SavePdf(Path.Combine(folder, "b.pdf"));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.pdf")), File.ReadAllBytes(Path.Combine(folder, "b.pdf")));
    }
}
