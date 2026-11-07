namespace Divit.Tests;

using Divit.Docx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Xunit;

/// <summary>
/// Streamed sections: content written while the document is saved, the same as stored content, in little memory.
/// </summary>
public class DocxStreamTests
{
    private static string Save(DocxDocument doc, string name, bool keep = false)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        doc.Save(path);
        Assert.Empty(DocxTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    private static DocxDocument NewDocument() =>
        new DocxDocument { Created = XlsxTestKit.FixedTime, NumberHeadings = true, Author = "Divit" };

    /// <summary>
    /// Content of every kind a streamed section takes, added to any container.
    /// </summary>
    private static void Fill(DocxContainer c, DocxSource source)
    {
        c.AddParagraph("Measurement log", null, DocxStyle.Heading(1));
        c.AddParagraph().Add("Bold ", new DocxFont { Bold = true }).Add("and a link: ").AddLink("divit.example", "https://divit.example")
            .AddFootnote("A streamed footnote.").Add(" ").AddCitation(source).Add(" ").AddCheckBox(true)
            .AddDeletion("old").AddInsertion("new").AddEquation(DocxMath.Sup("x", "2"));
        var steps = c.AddList(DocxListKind.Numbered);
        steps.AddItem("First step");
        steps.AddItem("Detail", 1);
        steps.AddItem("Second step");
        c.AddParagraph().AddImage(TestImages.Png(40, 20, (x, y) => 0x1F4E79), 2);

        var table = c.AddTable(3).SetColumnWidths(3, 5, 5);
        table.AddHeaderRow("Point", "Frequency", "Gain");
        for (var i = 1; i <= 6; i++)
        {
            var row = table.AddRow(i.ToString(CultureInfo.InvariantCulture), (8 + (i * 0.1)).ToString("0.0", CultureInfo.InvariantCulture) + " GHz");
            if (i == 2)
            {
                // A span over rows that are written before the rows it covers.
                row.Cells[2].RowSpan = 3;
                row.Cells[2].SetText("8 dBi");
                row.Cells[0].Shading = "#EEF5FB";
            }
            else if (i < 2 || i > 4)
            {
                row.Cells[2].SetText("8.1 dBi");
            }
        }

        c.AddParagraph("After the table.");
    }

    private static DocxSource AddSource(DocxDocument doc)
    {
        var source = doc.AddSource("IEEE149", DocxSourceKind.Report, "Antenna Measurements");
        source.Corporate = "IEEE";
        source.Year = "2021";
        return source;
    }

    /// <summary>
    /// The body with relationship ids taken out: a streamed section adds some relationships later than a stored
    /// one, which changes their ids but nothing else.
    /// </summary>
    private static string Body(string path)
    {
        using (var document = WordprocessingDocument.Open(path, false))
        {
            return Regex.Replace(document.MainDocumentPart!.Document!.Body!.OuterXml, "rId[0-9]+", "rId");
        }
    }

    [Fact]
    public void AStreamedSectionWritesWhatAStoredOneWrites()
    {
        var stored = NewDocument();
        stored.AddParagraph("Cover");
        var storedSource = AddSource(stored);
        Fill(stored.AddSection(), storedSource);

        var streamed = NewDocument();
        streamed.AddParagraph("Cover");
        var streamedSource = AddSource(streamed);
        streamed.AddStreamedSection(w => Fill(w, streamedSource));

        var storedBody = Body(Save(stored, "stored.docx"));
        var streamedPath = Save(streamed, "streamed.docx");
        Assert.Equal(storedBody, Body(streamedPath));

        using (var document = WordprocessingDocument.Open(streamedPath, false))
        {
            var main = document.MainDocumentPart!;
            Assert.NotNull(main.NumberingDefinitionsPart);
            Assert.NotNull(main.FootnotesPart);
            Assert.Single(main.ImageParts);
            Assert.Single(main.HyperlinkRelationships);
            Assert.Contains(main.StyleDefinitionsPart!.Styles!.Elements<Style>(), s => s.StyleId == DocxStyle.ListParagraph);
            Assert.Equal(2, main.Document!.Body!.Descendants<SectionProperties>().Count());
            var merged = main.Document.Body.Descendants<VerticalMerge>().ToList();
            Assert.Equal(3, merged.Count);
            Assert.Equal(MergedCellValues.Restart, merged[0].Val!.Value);
        }
    }

    [Fact]
    public void AStreamedDocumentIsValidAndKept()
    {
        var doc = NewDocument();
        doc.Look.BodyFont = "Trebuchet MS";
        doc.Title = "Streamed measurement log";
        doc.AddTitle("Streamed measurement log");
        doc.AddTableOfContents(2);
        var log = doc.AddStreamedSection(w =>
        {
            for (var band = 1; band <= 4; band++)
            {
                w.AddHeading("Band " + band.ToString(CultureInfo.InvariantCulture), 1);
                w.AddParagraph("Gain measured every 10 MHz; the table repeats its header on every page.");
                var table = w.AddTable(4).SetColumnWidths(2, 4, 4, 4);
                table.AlignColumn(2, DocxAlignment.Right).AlignColumn(3, DocxAlignment.Right).AlignColumn(4, DocxAlignment.Right);
                table.AddHeaderRow("No", "Frequency (GHz)", "Gain (dBi)", "VSWR");
                for (var i = 0; i < 250; i++)
                {
                    var f = 2 + (band * 7.5) + (i * 0.01);
                    table.AddRow(
                        (i + 1).ToString(CultureInfo.InvariantCulture),
                        f.ToString("0.00", CultureInfo.InvariantCulture),
                        (8 + Math.Sin(f)).ToString("0.00", CultureInfo.InvariantCulture),
                        (1.2 + (0.1 * Math.Cos(f))).ToString("0.00", CultureInfo.InvariantCulture));
                }

                w.AddPageBreak();
            }
        });
        log.Header.AddParagraph().Add("Streamed log · page ").AddField(DocxField.Page);
        log.Watermark = "LOG";
        Save(doc, "word-streamed-log.docx", keep: true);
    }

    [Fact]
    public void RowsAndParagraphsAreNotKept()
    {
        // Memory after the first thousand rows and after the last one, while the table is still open; the same for
        // paragraphs. Kept, 20 000 rows or paragraphs take several megabytes; streamed, memory stays where it was.
        long rowsEarly = 0, rowsLate = 0, paragraphsEarly = 0, paragraphsLate = 0;
        var doc = NewDocument();
        doc.AddStreamedSection(w =>
        {
            var table = w.AddTable(3);
            for (var i = 0; i < 20000; i++)
            {
                table.AddRow("Row " + i.ToString(CultureInfo.InvariantCulture), "8.2 GHz", "8 dBi");
                if (i == 1000) { rowsEarly = GC.GetTotalMemory(true); }
            }

            rowsLate = GC.GetTotalMemory(true);
            for (var i = 0; i < 20000; i++)
            {
                w.AddParagraph("Paragraph " + i.ToString(CultureInfo.InvariantCulture));
                if (i == 1000) { paragraphsEarly = GC.GetTotalMemory(true); }
            }

            paragraphsLate = GC.GetTotalMemory(true);
        });
        doc.Save(Path.Combine(XlsxTestKit.NewFolder(), "large.docx"));
        Assert.True(rowsLate - rowsEarly < 1024 * 1024, "rows grew memory by " + ((rowsLate - rowsEarly) / 1024).ToString(CultureInfo.InvariantCulture) + " KB");
        Assert.True(paragraphsLate - paragraphsEarly < 1024 * 1024, "paragraphs grew memory by " + ((paragraphsLate - paragraphsEarly) / 1024).ToString(CultureInfo.InvariantCulture) + " KB");
    }

    [Fact]
    public void ARowCanChangeUntilTheNextOne()
    {
        var doc = NewDocument();
        doc.AddStreamedSection(w =>
        {
            var table = w.AddTable(2);
            var row = table.AddRow("a", "b");
            row.Cells[1].Shading = "#FF0000";
            row.KeepTogether = true;
            table.AddRow("c", "d");
            var last = w.AddParagraph("Last");
            last.Alignment = DocxAlignment.Center;
        });
        var path = Save(doc, "changes.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            var rows = body.Descendants<TableRow>().ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal("FF0000", rows[0].Descendants<Shading>().Single().Fill!.Value);
            Assert.NotNull(rows[0].TableRowProperties!.GetFirstChild<CantSplit>());
            Assert.Equal(JustificationValues.Center, body.Elements<Paragraph>().Single(p => p.InnerText == "Last").ParagraphProperties!.Justification!.Val!.Value);
        }
    }

    [Fact]
    public void LinksAndReferencesToParagraphsAreRefused()
    {
        // A streamed paragraph exists only while a save runs: pointing at it later finds no bookmark.
        DocxParagraph? inside = null;
        var into = NewDocument();
        into.AddStreamedSection(w => inside = w.AddParagraph("Inside"));
        into.Save(Path.Combine(XlsxTestKit.NewFolder(), "first.docx"));
        into.AddParagraph().AddReference(inside!);
        Assert.Throws<InvalidOperationException>(() => into.Save(Path.Combine(XlsxTestKit.NewFolder(), "into.docx")));

        var outOf = NewDocument();
        var heading = outOf.AddHeading("Stored");
        outOf.AddStreamedSection(w => w.AddParagraph().AddReference(heading));
        Assert.Throws<InvalidOperationException>(() => outOf.Save(Path.Combine(XlsxTestKit.NewFolder(), "out.docx")));

        var link = NewDocument();
        var target = link.AddHeading("Stored");
        link.AddStreamedSection(w => w.AddParagraph().AddLink("back", target));
        Assert.Throws<InvalidOperationException>(() => link.Save(Path.Combine(XlsxTestKit.NewFolder(), "link.docx")));

        var foreign = new DocxDocument().AddSource("X", DocxSourceKind.Misc, "Elsewhere");
        var cite = NewDocument();
        cite.AddStreamedSection(w => w.AddParagraph().AddCitation(foreign));
        Assert.Throws<InvalidOperationException>(() => cite.Save(Path.Combine(XlsxTestKit.NewFolder(), "cite.docx")));
    }

    [Fact]
    public void EveryFeatureOnlyInAStreamIsValid()
    {
        // Without numbered headings, so the numbering part and its relationship come from the stream alone.
        var doc = new DocxDocument { Created = XlsxTestKit.FixedTime };
        doc.AddStreamedSection(w =>
        {
            w.AddList().AddItem("Bullet");
            w.AddParagraph("Note").AddFootnote("Only footnote.");
            w.AddEquation(DocxMath.Fraction("a", "b"));
            w.AddParagraph().AddCheckBox();
        });
        var path = Save(doc, "stream-only.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            Assert.NotNull(main.NumberingDefinitionsPart);
            Assert.Equal(3, main.FootnotesPart!.Footnotes!.Elements<Footnote>().Count());
            Assert.Single(main.NumberingDefinitionsPart!.Numbering!.Elements<NumberingInstance>());
            var root = main.Document!.OuterXml.Substring(0, main.Document.OuterXml.IndexOf("<w:body", StringComparison.Ordinal));
            Assert.Contains("mc:Ignorable=\"w14\"", root);
            Assert.Contains("xmlns:m=", root);
        }
    }

    [Fact]
    public void StreamedDocumentsAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        var doc = NewDocument();
        var source = AddSource(doc);
        doc.AddStreamedSection(w => Fill(w, source));
        doc.Save(Path.Combine(folder, "a.docx"));
        doc.Save(Path.Combine(folder, "b.docx"));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.docx")), File.ReadAllBytes(Path.Combine(folder, "b.docx")));
    }
}
