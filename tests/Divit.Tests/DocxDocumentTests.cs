namespace Divit.Tests;

using Divit.Docx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using Xunit;

/// <summary>
/// The Word writer's foundation: sections and page setup, paragraphs and fonts, numbered headings, the house style,
/// headers and footers with page fields, and the table of contents.
/// </summary>
public class DocxDocumentTests
{
    private static readonly DocxFont Bold = new DocxFont { Bold = true };

    /// <summary>
    /// The opening pages of a test procedure: a cover with a letterhead footer, the contents, numbered chapters.
    /// </summary>
    private static DocxDocument Procedure()
    {
        var doc = new DocxDocument { Title = "Antenna Test Procedure", Author = "Divit", Created = XlsxTestKit.FixedTime, Language = "en-GB", NumberHeadings = true };
        doc.Look.BodyFont = "Trebuchet MS";
        doc.Look.BodySize = 12;
        doc.Look.HeadingColor = "#1F4E79";

        // Cover: centred on the page, its own footer as a wide letterhead, no page number.
        var cover = doc.Section;
        cover.VerticalAlignment = DocxVerticalAlignment.Center;
        doc.AddTitle("Antenna Test Procedure").Alignment = DocxAlignment.Center;
        doc.AddSubtitle("X-band patch antenna, part no. ANT-1234").Alignment = DocxAlignment.Center;
        var info = doc.AddParagraph();
        info.Alignment = DocxAlignment.Center;
        info.Add("Document no: ", Bold).Add("TP-ANT-001\n").Add("Revision: ", Bold).Add("A\n").Add("Date: ", Bold).AddField(DocxField.Date);
        var letterhead = cover.FirstPageFooter.AddParagraph();
        letterhead.Alignment = DocxAlignment.Center;
        letterhead.Shading = "#1F4E79";
        letterhead.Add("DIVIT ELECTRONICS  ·  Test and Qualification Department  ·  Ankara  ·  divit.example", new DocxFont { Color = "#FFFFFF", Size = 9 });

        // The body: header with document number and page x of y, footer with the title.
        var body = doc.AddSection();
        body.VerticalAlignment = DocxVerticalAlignment.Top;
        var header = body.Header.AddParagraph();
        header.AddTabStop(16, DocxTabAlignment.Right);
        header.Border = new DocxBorder("#9DC3E6");
        header.Add("TP-ANT-001  Rev. A", new DocxFont { Size = 9, Color = "#1F4E79" }).AddTab()
            .Add("Page ", new DocxFont { Size = 9 }).AddField(DocxField.Page, new DocxFont { Size = 9 })
            .Add(" / ", new DocxFont { Size = 9 }).AddField(DocxField.NumPages, new DocxFont { Size = 9 });
        body.Footer.AddParagraph().Add("Antenna Test Procedure", new DocxFont { Size = 9, Italic = true }).Alignment = DocxAlignment.Center;

        doc.AddTableOfContents(3, "Contents");
        doc.AddPageBreak();

        doc.AddHeading("Scope", 1);
        doc.AddParagraph("This procedure describes how the X-band patch antenna is tested before delivery: the equipment, the set-up, the software and the steps.");
        doc.AddHeading("General information", 1);
        doc.AddHeading("Project", 2);
        doc.AddParagraph().Add("Project: ", Bold).Add("Divit demo satellite. ").Add("Stock no: ", Bold).Add("1234-56-789-0000.");
        doc.AddHeading("Unit under test", 2);
        doc.AddParagraph().Add("Gain 8 dBi, VSWR ≤ 1.5, impedance 50 Ω, frequency 8.0–8.4 GHz; the wavelength λ").Add("0", new DocxFont { Script = DocxScript.Subscript }).Add(" is 3.6 cm.");
        doc.AddHeading("Tests", 1);
        doc.AddHeading("Return loss test", 2);
        doc.AddHeading("Test equipment", 3);
        doc.AddParagraph("The equipment list follows in a table.");
        doc.AddHeading("Test set-up", 3);
        doc.AddParagraph("The antenna is connected to port 1 of the calibrated network analyser.");
        doc.AddHeading("Software preparation", 3);
        var note = doc.AddParagraph();
        note.Shading = "#EEF5FB";
        note.Border = new DocxBorder("#2E75B6", 1);
        note.Add("Note: ", new DocxFont { Bold = true, Color = "#1F4E79" }).Add("install the analyser software before connecting the instrument.");
        doc.AddHeading("Radiation pattern test", 2);
        doc.AddParagraph("Measured in the anechoic chamber.");
        return doc;
    }

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

    [Fact]
    public void AProcedureIsValid()
    {
        var path = Save(Procedure(), "procedure.docx", keep: true);
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            Assert.Single(main.HeaderParts);
            Assert.Equal(2, main.FooterParts.Count());
            Assert.NotNull(main.StyleDefinitionsPart);
            Assert.NotNull(main.NumberingDefinitionsPart);
            Assert.NotNull(main.ThemePart);
            Assert.Equal("Antenna Test Procedure", document.PackageProperties.Title);
        }
    }

    [Fact]
    public void SectionsCarryTheirPageSetup()
    {
        var path = Save(Procedure(), "sections-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document.Body!;
            var sections = body.Descendants<SectionProperties>().ToList();
            Assert.Equal(2, sections.Count);

            // The cover: centred, first page different, its footer the letterhead.
            var cover = sections[0];
            Assert.Equal(VerticalJustificationValues.Center, cover.GetFirstChild<VerticalTextAlignmentOnPage>()!.Val!.Value);
            Assert.NotNull(cover.GetFirstChild<TitlePage>());
            Assert.Equal(HeaderFooterValues.First, cover.Elements<FooterReference>().Single().Type!.Value);
            var size = cover.GetFirstChild<PageSize>()!;
            Assert.Equal((11906u, 16838u), (size.Width!.Value, size.Height!.Value));
            Assert.Equal(1417, cover.GetFirstChild<PageMargin>()!.Top!.Value);

            // The body: its own header and footer, top aligned, and it is the body's last element.
            var last = sections[1];
            Assert.Same(body.LastChild, last);
            Assert.Null(last.GetFirstChild<TitlePage>());
            Assert.Equal(HeaderFooterValues.Default, last.Elements<HeaderReference>().Single().Type!.Value);
            Assert.Null(last.GetFirstChild<VerticalTextAlignmentOnPage>());
        }
    }

    [Fact]
    public void HeadersHoldFieldsAndTabs()
    {
        var path = Save(Procedure(), "header-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var body = main.Document.Body!.Descendants<SectionProperties>().Last();
            var header = (HeaderPart)main.GetPartById(body.Elements<HeaderReference>().Single().Id!.Value!);
            var fields = header.Header.Descendants<SimpleField>().Select(f => f.Instruction!.Value).ToArray();
            Assert.Equal(new[] { " PAGE ", " NUMPAGES " }, fields);
            var tab = header.Header.Descendants<TabStop>().Single();
            Assert.Equal(TabStopValues.Right, tab.Val!.Value);
            Assert.Equal(9071, tab.Position!.Value);
            Assert.NotNull(header.Header.Descendants<ParagraphBorders>().Single().TopBorder);

            var date = main.Document.Body!.Descendants<SimpleField>().Single();
            Assert.Equal(" DATE \\@ \"dd.MM.yyyy\" ", date.Instruction!.Value);
            Assert.Equal("02.01.2026", date.InnerText);
        }
    }

    [Fact]
    public void HeadingsAreNumberedAndListedInTheContents()
    {
        var path = Save(Procedure(), "toc-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;

            // One outline of six levels: 1, 1.1, 1.1.1 …, tied to the heading styles.
            var levels = main.NumberingDefinitionsPart!.Numbering.Descendants<Level>().ToList();
            Assert.Equal(6, levels.Count);
            Assert.Equal(new[] { "%1", "%1.%2", "%1.%2.%3" }, levels.Take(3).Select(l => l.LevelText!.Val!.Value).ToArray());
            Assert.Equal("Heading3", levels[2].ParagraphStyleIdInLevel!.Val!.Value);
            var heading2 = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(s => s.StyleId == "Heading2");
            Assert.Equal(1, heading2.StyleParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value);
            Assert.Equal(1, heading2.StyleParagraphProperties.OutlineLevel!.Val!.Value);

            // The contents: a TOC field around entries 1 … 3.2 for levels one to three.
            var body = main.Document.Body!;
            var instruction = body.Descendants<FieldCode>().Single();
            Assert.Equal(" TOC \\o \"1-3\" \\h \\z \\u ", instruction.Text);
            var entries = body.Elements<Paragraph>()
                .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("TOC", StringComparison.Ordinal) == true
                    && p.ParagraphProperties.ParagraphStyleId.Val.Value != "TOCHeading")
                .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToArray();
            Assert.Equal(new[] { "1 Scope", "2 General information", "2.1 Project", "2.2 Unit under test", "3 Tests", "3.1 Return loss test",
                "3.1.1 Test equipment", "3.1.2 Test set-up", "3.1.3 Software preparation", "3.2 Radiation pattern test" }, entries);
            Assert.Equal(new[] { "begin", "separate", "end" },
                body.Descendants<FieldChar>().Select(f => f.FieldCharType!.InnerText).ToArray());
            Assert.True(main.DocumentSettingsPart!.Settings.GetFirstChild<UpdateFieldsOnOpen>()!.Val!.Value);

            // The contents title is not numbered and not listed.
            var tocHeading = main.StyleDefinitionsPart.Styles.Elements<Style>().Single(s => s.StyleId == "TOCHeading");
            Assert.Equal(0, tocHeading.StyleParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);
            Assert.Equal(9, tocHeading.StyleParagraphProperties.OutlineLevel!.Val!.Value);
        }
    }

    [Fact]
    public void TheLookBecomesTheStyles()
    {
        var path = Save(Procedure(), "styles-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            var defaults = styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
            Assert.Equal("Trebuchet MS", defaults.RunFonts!.Ascii!.Value);
            Assert.Equal("24", defaults.FontSize!.Val!.Value);
            Assert.Equal("en-GB", defaults.Languages!.Val!.Value);
            var heading1 = styles.Elements<Style>().Single(s => s.StyleId == "Heading1");
            Assert.Equal("heading 1", heading1.StyleName!.Val!.Value);
            Assert.Equal("1F4E79", heading1.StyleRunProperties!.Color!.Val!.Value);
            Assert.Equal("32", heading1.StyleRunProperties.FontSize!.Val!.Value);
            Assert.NotNull(heading1.StyleRunProperties.Bold);
            Assert.Null(styles.Elements<Style>().Single(s => s.StyleId == "Heading4").StyleRunProperties!.Bold);
            Assert.Contains(styles.Elements<Style>(), s => s.StyleId == "TableNormal" && s.Default!.Value);

            var compat = document.MainDocumentPart.DocumentSettingsPart!.Settings.Descendants<CompatibilitySetting>().Single();
            Assert.Equal("15", compat.Val!.Value);
        }
    }

    [Fact]
    public void RunsKeepTheirFonts()
    {
        var doc = new DocxDocument();
        doc.AddParagraph().Add("plain ").Add("bold", Bold).Add(" x").Add("2", new DocxFont { Script = DocxScript.Superscript })
            .Add(" red", new DocxFont { Color = "#C00000", Size = 14, Name = "Consolas", Italic = true, Underline = true, Strikethrough = true, AllCaps = true, Background = "#FFFF00" });
        doc.AddParagraph("line one\nline two\tafter a tab");
        var path = Save(doc, "runs-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var paragraphs = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();
            var runs = paragraphs[0].Elements<Run>().ToList();
            Assert.Null(runs[0].RunProperties);
            Assert.Equal(" ", runs[0].InnerText.Substring(5));
            Assert.Equal(DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve, runs[0].GetFirstChild<Text>()!.Space!.Value);
            Assert.NotNull(runs[1].RunProperties!.Bold);
            Assert.Equal(VerticalPositionValues.Superscript, runs[3].RunProperties!.VerticalTextAlignment!.Val!.Value);
            var red = runs[4].RunProperties!;
            Assert.Equal(("C00000", "28", "Consolas"), (red.Color!.Val!.Value, red.FontSize!.Val!.Value, red.RunFonts!.Ascii!.Value));
            Assert.NotNull(red.Italic);
            Assert.NotNull(red.Strike);
            Assert.NotNull(red.Caps);
            Assert.Equal(UnderlineValues.Single, red.Underline!.Val!.Value);
            Assert.Equal("FFFF00", red.Shading!.Fill!.Value);

            var broken = paragraphs[1].Elements<Run>().ToList();
            Assert.Equal(5, broken.Count);
            Assert.NotNull(broken[1].GetFirstChild<Break>());
            Assert.NotNull(broken[3].GetFirstChild<TabChar>());
        }
    }

    [Fact]
    public void ParagraphLayoutIsWritten()
    {
        var doc = new DocxDocument();
        var p = doc.AddParagraph("laid out");
        p.Alignment = DocxAlignment.Justify;
        p.SpaceBefore = 12;
        p.SpaceAfter = 0;
        p.LineSpacing = 1.5;
        p.LeftIndent = 1;
        p.FirstLineIndent = -0.5;
        p.KeepWithNext = true;
        p.PageBreakBefore = true;
        doc.AddPageBreak();
        var landscape = doc.AddSection();
        landscape.Landscape = true;
        landscape.PageSize = DocxPageSize.A3;
        landscape.PageNumberStart = 1;
        landscape.PageNumberFormat = DocxNumberFormat.LowerRoman;
        landscape.Margins = new DocxMargins(2, 2, 1.5, 1.5, 1, 1);
        doc.AddParagraph("wide");
        var path = Save(doc, "layout-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document.Body!;
            var props = body.Elements<Paragraph>().First().ParagraphProperties!;
            Assert.Equal(JustificationValues.Both, props.Justification!.Val!.Value);
            Assert.Equal(("240", "0", "360"), (props.SpacingBetweenLines!.Before!.Value, props.SpacingBetweenLines.After!.Value, props.SpacingBetweenLines.Line!.Value));
            Assert.Equal(("567", "283"), (props.Indentation!.Left!.Value, props.Indentation.Hanging!.Value));
            Assert.NotNull(props.KeepNext);
            Assert.NotNull(props.PageBreakBefore);
            Assert.Equal(BreakValues.Page, body.Descendants<Break>().Single().Type!.Value);

            var wide = body.Descendants<SectionProperties>().Last();
            var size = wide.GetFirstChild<PageSize>()!;
            Assert.Equal((23811u, 16838u, PageOrientationValues.Landscape), (size.Width!.Value, size.Height!.Value, size.Orient!.Value));
            var numbers = wide.GetFirstChild<PageNumberType>()!;
            Assert.Equal((NumberFormatValues.LowerRoman, 1), (numbers.Format!.Value, numbers.Start!.Value));
            Assert.Equal(850u, wide.GetFirstChild<PageMargin>()!.Left!.Value);

            // The new section took the previous one's settings until changed.
            var first = body.Descendants<SectionProperties>().First();
            Assert.Equal(11906u, first.GetFirstChild<PageSize>()!.Width!.Value);
        }
    }

    [Fact]
    public void DocumentsAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        Procedure().Save(Path.Combine(folder, "a.docx"));
        Procedure().Save(Path.Combine(folder, "b.docx"));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.docx")), File.ReadAllBytes(Path.Combine(folder, "b.docx")));
    }

    [Fact]
    public void AnEmptyDocumentIsValid()
    {
        var path = Save(new DocxDocument(), "empty-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            Assert.Null(document.MainDocumentPart!.NumberingDefinitionsPart);
            Assert.Null(document.MainDocumentPart.DocumentSettingsPart!.Settings.GetFirstChild<UpdateFieldsOnOpen>());
        }
    }
}
