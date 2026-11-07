namespace Divit.Tests;

using Divit.Docx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using Xunit;

/// <summary>
/// A full acceptance test procedure for the SH2000 horn antenna, built only from the library's general blocks:
/// cover with letterhead, headers and footers as tables, change record, contents, lists of figures and tables,
/// numbered chapters, banded tables, figures with captions, step tables and a record form.
/// </summary>
public class DocxProcedureTests
{
    private static readonly DocxFont Bold = new DocxFont { Bold = true };
    private static readonly DocxFont Small = new DocxFont { Size = 8 };
    private static readonly DocxFont SmallBold = new DocxFont { Size = 8, Bold = true, Color = "#1F4E79" };
    private static readonly DocxFont White = new DocxFont { Size = 9, Bold = true, Color = "#FFFFFF" };

    private const string DocumentNo = "ATP-SH2000-001";
    private const string StockNo = "5985-01-234-5678";
    private const string Revision = "B";
    private const string ReleaseDate = "07.11.2026";

    internal static DocxDocument Procedure()
    {
        var doc = new DocxDocument
        {
            Title = "SH2000 Acceptance Test Procedure",
            Subject = "Dual-ridge horn antenna, 2–32 GHz",
            Author = "Test and Qualification Department",
            Keywords = DocumentNo,
            Created = XlsxTestKit.FixedTime,
            Language = "en-GB",
            NumberHeadings = true,
        };
        doc.Look.BodyFont = "Trebuchet MS";
        doc.Look.BodySize = 11;
        doc.Look.HeadingColor = "#1F4E79";
        doc.Look.HeadingSizes = new[] { 15.0, 13, 12, 11 };
        doc.Look.TableHeaderFill = "#BDD7EE";
        doc.Look.TableBandFill = "#EEF5FB";
        var logo = TestImages.Logo();

        Cover(doc, logo);
        Body(doc, logo);

        doc.AddHeading("Document change record", 1).Style = "TOCHeading";
        doc.AddCaption(DocxCaptionKind.Table, "Document change record");
        var changes = doc.AddTable(5).SetColumnWidths(1.5, 2.5, 7, 2.5, 2.5);
        changes.AddHeaderRow("Rev.", "Date", "Description", "Prepared", "Approved");
        changes.AddRow("A", "12.05.2026", "First issue", "E. Kaya", "M. Demir");
        changes.AddRow("B", ReleaseDate, "Software installation steps and record form added", "E. Kaya", "M. Demir");
        changes.AlignColumn(1, DocxAlignment.Center);

        doc.AddPageBreak();
        doc.AddTableOfContents(3, "Contents");
        doc.AddTableOfFigures(DocxCaptionKind.Figure, "List of figures");
        doc.AddTableOfFigures(DocxCaptionKind.Table, "List of tables");
        doc.AddPageBreak();

        Scope(doc);
        General(doc);
        Documents(doc);
        Abbreviations(doc);
        Tests(doc);
        Annex(doc);
        return doc;
    }

    /// <summary>
    /// The cover: logo, title block and document data, with the wide letterhead of signatures as the first page's
    /// footer.
    /// </summary>
    private static void Cover(DocxDocument doc, byte[] logo)
    {
        var cover = doc.Section;
        cover.Margins = new DocxMargins(2, 2, 2, 2, 1, 0.8);
        doc.AddParagraph().AddImage(logo, 6, 0, "Divit Electronics logo").Alignment = DocxAlignment.Left;
        doc.AddParagraph().SpaceBefore = 110;
        doc.AddTitle("Acceptance Test Procedure").Alignment = DocxAlignment.Center;
        doc.AddSubtitle("SH2000 Dual-Ridge Horn Antenna, 2–32 GHz").Alignment = DocxAlignment.Center;
        var data = doc.AddTable(2).SetColumnWidths(5, 7);
        data.Look = DocxTableLook.Grid;
        data.Centered = true;
        foreach (var (label, value) in new[] { ("Document no", DocumentNo), ("Stock no", StockNo), ("Revision", Revision), ("Release date", ReleaseDate), ("Classification", "Company confidential") })
        {
            var row = data.AddRow();
            row[1].SetText(label, Bold);
            row[1].Shading = "#EEF5FB";
            row[2].SetText(value);
        }

        // The letterhead: who prepared, checked and approved, and when the document was released.
        var footer = cover.FirstPageFooter.AddTable(4);
        footer.Look = DocxTableLook.Grid;
        var heads = footer.AddRow();
        var titles = new[] { "Prepared by", "Checked by", "Approved by", "Release" };
        for (var i = 0; i < titles.Length; i++)
        {
            heads.Cells[i].Shading = "#1F4E79";
            heads.Cells[i].SetText(titles[i], White).Alignment = DocxAlignment.Center;
        }

        var names = footer.AddRow();
        names[1].SetText("E. Kaya\nTest Engineer", Small);
        names[2].SetText("A. Yılmaz\nQuality Assurance", Small);
        names[3].SetText("M. Demir\nProgramme Manager", Small);
        names[4].SetText("Date: " + ReleaseDate + "\nRevision: " + Revision, Small);
        names.Height = 1.6;
        var signs = footer.AddRow();
        foreach (var (cell, text) in signs.Cells.Zip(new[] { "Signature:", "Signature:", "Signature:", "Transfer date: " + ReleaseDate }, (c, t) => (c, t)))
        {
            cell.SetText(text, Small);
        }
        var company = footer.AddRow();
        company[1].ColumnSpan = 4;
        company[1].Shading = "#DEEAF6";
        company[1].SetText("DIVIT ELECTRONICS  ·  Test and Qualification Department  ·  Ankara, Türkiye  ·  divit.example  ·  This document is the property of Divit Electronics and may not be copied or passed on without its written consent.", Small).Alignment = DocxAlignment.Center;
    }

    /// <summary>
    /// Header and footer of the body pages: logo, title and identifiers above; release data and page below.
    /// </summary>
    private static void Body(DocxDocument doc, byte[] logo)
    {
        var body = doc.AddSection();
        body.Margins = new DocxMargins(2.5, 2.2, 2, 2, 1, 1);
        var header = body.Header.AddTable(3).SetColumnWidths(4, 7.5, 5.5);
        header.Look = DocxTableLook.Grid;
        var top = header.AddRow();
        top[1].AddParagraph().AddImage(logo, 3.2);
        top[1].VerticalAlignment = DocxCellAlignment.Center;
        top[2].SetText("SH2000 Acceptance Test Procedure", new DocxFont { Bold = true, Color = "#1F4E79", Size = 10 }).Alignment = DocxAlignment.Center;
        top[2].VerticalAlignment = DocxCellAlignment.Center;
        var ids = top[3].SetText("Doc. no: ", SmallBold);
        ids.Add(DocumentNo + "\n", Small).Add("Stock no: ", SmallBold).Add(StockNo + "\n", Small).Add("Revision: ", SmallBold).Add(Revision, Small);
        body.Header.AddParagraph();

        var footer = body.Footer.AddTable(3).SetColumnWidths(5.5, 6, 5.5);
        footer.Look = DocxTableLook.Grid;
        var row = footer.AddRow();
        row[1].SetText("Release date: ", SmallBold).Add(ReleaseDate, Small);
        row[2].SetText("Revision: ", SmallBold).Add(Revision, Small).Alignment = DocxAlignment.Center;
        var page = row[3].SetText("Page ", SmallBold);
        page.AddField(DocxField.Page, Small).Add(" / ", Small).AddField(DocxField.NumPages, Small).Alignment = DocxAlignment.Right;
        foreach (var cell in row.Cells)
        {
            cell.Shading = "#EEF5FB";
        }

        body.Footer.AddParagraph("Company confidential – printed copies are not controlled.", new DocxFont { Size = 7, Italic = true, Color = "#7F7F7F" }).Alignment = DocxAlignment.Center;
    }

    private static void Scope(DocxDocument doc)
    {
        doc.AddHeading("Scope", 1);
        doc.AddParagraph("This procedure describes how the SH2000 dual-ridge horn antenna is tested before it is accepted: the hardware and software needed, how the test set-up is built, how the software is prepared and installed, and the steps of each test. Measured values are recorded on the form in Annex A, not in this document.");
        doc.AddParagraph("The procedure applies to every SH2000 delivered under the project, at the acceptance test after production.");
    }

    private static void General(DocxDocument doc)
    {
        doc.AddHeading("General information", 1);
        doc.AddHeading("Project", 2);
        doc.AddParagraph().Add("The SH2000 is used as the reference antenna of the ").Add("Divit EMC Range", Bold).Add(" project, where it measures the gain and pattern of the antennas under development.");
        doc.AddHeading("Stock definition", 2);
        doc.AddCaption(DocxCaptionKind.Table, "Stock definition");
        var stock = doc.AddTable(2).SetColumnWidths(5, 12);
        stock.AddHeaderRow("Item", "Value");
        foreach (var (item, value) in new[] { ("Stock no", StockNo), ("Part no", "617-355-SH2000"), ("Manufacturer", "Microwave Vision Group"), ("Description", "Dual-ridge horn antenna, 2–32 GHz"), ("Unit of issue", "Each") })
        {
            stock.AddRow(item, value);
        }

        doc.AddHeading("Unit introduction", 2);
        doc.AddParagraph("The SH2000 is a broadband dual-ridge horn. The ridges keep the impedance near 50 Ω over the whole band, so one antenna covers 2–32 GHz. It is fed through an SMA connector at the back of the waveguide section.");
        var figure = doc.AddParagraph();
        figure.Alignment = DocxAlignment.Center;
        figure.KeepWithNext = true;
        figure.AddImage(TestImages.Antenna(), 11, 0, "SH2000 seen from the side");
        doc.AddCaption(DocxCaptionKind.Figure, "SH2000 dual-ridge horn antenna").Alignment = DocxAlignment.Center;
        doc.AddCaption(DocxCaptionKind.Table, "Main characteristics");
        var specs = doc.AddTable(3).SetColumnWidths(6, 6, 5);
        specs.AddHeaderRow("Characteristic", "Requirement", "Verified in");
        specs.AddRow("Frequency range", "2–32 GHz", "5.1");
        specs.AddRow("VSWR", "≤ 2.0 : 1", "5.1");
        specs.AddRow("Gain", "3–17 dBi", "5.2");
        specs.AddRow("Mass", "≤ 0.3 kg", "Visual inspection");
        specs.AlignColumn(3, DocxAlignment.Center);
    }

    private static void Documents(DocxDocument doc)
    {
        doc.AddHeading("Documents", 1);
        doc.AddHeading("Applicable documents", 2);
        doc.AddParagraph("The following documents form part of this procedure; their latest revision applies.");
        doc.AddCaption(DocxCaptionKind.Table, "Applicable documents");
        var applicable = doc.AddTable(4).SetColumnWidths(1.5, 4.5, 9, 2);
        applicable.AddHeaderRow("No", "Document no", "Title", "Rev.");
        applicable.AddRow("AD1", "SPEC-SH2000-01", "SH2000 product specification", "C");
        applicable.AddRow("AD2", "QP-TEST-004", "Calibration of test equipment", "F");
        applicable.AddRow("AD3", "IEEE Std 149-2021", "Recommended practice for antenna measurements", "–");
        applicable.AlignColumn(1, DocxAlignment.Center).AlignColumn(4, DocxAlignment.Center);
        doc.AddHeading("Reference documents", 2);
        doc.AddCaption(DocxCaptionKind.Table, "Reference documents");
        var reference = doc.AddTable(4).SetColumnWidths(1.5, 4.5, 9, 2);
        reference.AddHeaderRow("No", "Document no", "Title", "Rev.");
        reference.AddRow("RD1", "UM-VNA-220", "Vector network analyser user manual", "3");
        reference.AlignColumn(1, DocxAlignment.Center).AlignColumn(4, DocxAlignment.Center);
    }

    private static void Abbreviations(DocxDocument doc)
    {
        doc.AddHeading("Abbreviations and symbols", 1);
        doc.AddCaption(DocxCaptionKind.Table, "Abbreviations");
        var abbreviations = doc.AddTable(2).SetColumnWidths(3.5, 13.5);
        abbreviations.AddHeaderRow("Abbreviation", "Meaning");
        foreach (var (a, m) in new[] { ("ATP", "Acceptance test procedure"), ("AUT", "Antenna under test"), ("DUT", "Device under test"), ("SMA", "SubMiniature version A connector"), ("VNA", "Vector network analyser"), ("VSWR", "Voltage standing wave ratio") })
        {
            abbreviations.AddRow(a, m);
        }

        doc.AddCaption(DocxCaptionKind.Table, "Symbols");
        var symbols = doc.AddTable(3).SetColumnWidths(2.5, 10.5, 4);
        symbols.AddHeaderRow("Symbol", "Quantity", "Unit");
        symbols.AddRow("f", "Frequency", "GHz");
        symbols.AddRow("G", "Gain", "dBi");
        var gamma = symbols.AddRow(null, "Reflection coefficient", "–");
        gamma[1].SetText("Γ");
        var lambda = symbols.AddRow(null, "Free-space wavelength", "m");
        lambda[1].SetText("λ").Add("0", new DocxFont { Script = DocxScript.Subscript });
        symbols.AlignColumn(1, DocxAlignment.Center).AlignColumn(3, DocxAlignment.Center);
    }

    private static void Tests(DocxDocument doc)
    {
        doc.AddHeading("Tests", 1);
        doc.AddHeading("Return loss (VSWR) test", 2);
        doc.AddParagraph("The test checks that the antenna is matched to 50 Ω over 2–32 GHz: the VSWR may not exceed 2.0 : 1 at any frequency.");

        doc.AddHeading("Hardware list", 3);
        doc.AddCaption(DocxCaptionKind.Table, "Hardware list");
        var hardware = doc.AddTable(6).SetColumnWidths(1, 4.5, 3.2, 2.8, 2.6, 2.9);
        hardware.AddHeaderRow("No", "Equipment", "Manufacturer", "Model", "Serial no", "Calibration due");
        hardware.AddRow("1", "Vector network analyser", "Rohde & Schwarz", "ZNA43", "101234", "03.2027");
        hardware.AddRow("2", "Calibration kit, 2.92 mm", "Rohde & Schwarz", "ZN-Z229", "100987", "01.2027");
        hardware.AddRow("3", "Test cable, 1 m", "Huber+Suhner", "SF102", "C-0441", "–");
        hardware.AddRow("4", "Torque wrench, 0.9 N·m", "Rohde & Schwarz", "ZN-ZTW", "T-017", "06.2027");
        hardware.AddRow("5", "Absorber panel", "Divit", "ABS-60", "–", "–");
        hardware.AlignColumn(1, DocxAlignment.Center);

        doc.AddHeading("Software list", 3);
        doc.AddCaption(DocxCaptionKind.Table, "Software list");
        var software = doc.AddTable(4).SetColumnWidths(1, 5, 2.5, 8.5);
        software.AddHeaderRow("No", "Software", "Version", "Purpose");
        software.AddRow("1", "VNA firmware", "2.40", "Runs the network analyser");
        software.AddRow("2", "Divit Antenna Test Suite", "1.3.0", "Drives the measurement and writes the record file");
        software.AddRow("3", "Calibration kit data", "ZN-Z229 V2", "Standards of the calibration kit");
        software.AlignColumn(1, DocxAlignment.Center).AlignColumn(3, DocxAlignment.Center);

        doc.AddHeading("Test set-up", 3);
        doc.AddParagraph("The antenna points at the absorber panel at least 1 m away, so the reflections of the room stay below the measured return loss (Figure 2).");
        var setUp = doc.AddParagraph();
        setUp.Alignment = DocxAlignment.Center;
        setUp.KeepWithNext = true;
        setUp.AddImage(TestImages.SetUp(), 15, 0, "VNA connected to the SH2000 through a calibrated cable, facing an absorber");
        doc.AddCaption(DocxCaptionKind.Figure, "Return loss test set-up").Alignment = DocxAlignment.Center;

        doc.AddHeading("Software preparation and installation", 3);
        var note = doc.AddParagraph();
        note.Shading = "#FFF4E5";
        note.Border = new DocxBorder("#ED7D31", 1);
        note.Add("Caution: ", new DocxFont { Bold = true, Color = "#C55A11" }).Add("install the software with the analyser switched off; a running analyser may lose its calibration.");
        doc.AddCaption(DocxCaptionKind.Table, "Software installation steps");
        Steps(doc, new[]
        {
            ("Copy the Divit Antenna Test Suite 1.3.0 installer to the test computer.", "Installer present"),
            ("Run the installer as administrator and accept the default folder.", "Installation finished without error"),
            ("Load the ZN-Z229 calibration kit data into the analyser.", "Kit listed under Calibration > Kits"),
            ("Start the test suite and select the SH2000 test plan.", "Plan opens, 2–32 GHz, 1601 points"),
        });

        doc.AddHeading("Test steps", 3);
        doc.AddCaption(DocxCaptionKind.Table, "Return loss test steps");
        Steps(doc, new[]
        {
            ("Switch the analyser on and let it warm up for 30 minutes.", "Ready indication"),
            ("Calibrate port 1 at the end of the test cable (open, short, load).", "Calibration valid"),
            ("Connect the SH2000 with the torque wrench at 0.9 N·m.", "Connector tight, not over-torqued"),
            ("Start the sweep from the test suite.", "VSWR curve shown"),
            ("Read the highest VSWR and record it on form A-1.", "VSWR ≤ 2.0 : 1"),
        });

        doc.AddHeading("Radiation pattern test", 2);
        doc.AddParagraph("The pattern is measured in the anechoic chamber with the positioner; the hardware, set-up and steps follow the same layout as 5.1 and are given in the chamber procedure (RD1).");
    }

    /// <summary>
    /// A table of steps: number, action, expected result, and a box to tick.
    /// </summary>
    private static void Steps(DocxDocument doc, (string Action, string Expected)[] steps)
    {
        var table = doc.AddTable(4).SetColumnWidths(1.2, 9, 5.3, 1.5);
        table.AddHeaderRow("Step", "Action", "Expected result", "✓");
        for (var i = 0; i < steps.Length; i++)
        {
            table.AddRow((i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), steps[i].Action, steps[i].Expected, "☐");
        }

        table.AlignColumn(1, DocxAlignment.Center).AlignColumn(4, DocxAlignment.Center);
    }

    private static void Annex(DocxDocument doc)
    {
        doc.AddPageBreak();
        doc.AddHeading("Annexes", 1);
        doc.AddHeading("Annex A – Test record form", 2);
        doc.AddParagraph("Fill in one form per antenna; the form is kept with the delivery documents.");
        doc.AddCaption(DocxCaptionKind.Table, "Form A-1: return loss test record");
        var form = doc.AddTable(4).SetColumnWidths(4.5, 4.5, 4, 4);
        form.Look = DocxTableLook.Grid;
        var head = form.AddHeaderRow("Antenna serial no", null, "Test date", null);
        head[2].Shading = "#FFFFFF";
        head[4].Shading = "#FFFFFF";
        form.AddRow("Tested by", null, "Test station", null);
        var section = form.AddRow("Results");
        section[1].ColumnSpan = 4;
        section[1].Shading = "#BDD7EE";
        form.AddRow("Frequency band", "Highest VSWR", "Limit", "Result");
        foreach (var band in new[] { "2–8 GHz", "8–18 GHz", "18–32 GHz" })
        {
            form.AddRow(band, null, "≤ 2.0 : 1", "☐ Pass   ☐ Fail");
        }

        var remarks = form.AddRow("Remarks");
        remarks[2].ColumnSpan = 3;
        remarks.Height = 2.5;
        var sign = form.AddRow("Signature", null, "Quality assurance", null);
        sign.Height = 1.5;
        var verdict = form.AddRow("Verdict", "☐ Accepted   ☐ Rejected");
        verdict[2].ColumnSpan = 3;
    }

    private static string Save(DocxDocument doc, string name)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        doc.Save(path);
        Assert.Empty(DocxTestKit.Validate(path));
        return path;
    }

    [Fact]
    public void TheProcedureIsValid()
    {
        var path = Save(Procedure(), "acceptance-test-procedure.docx");
#if NETFRAMEWORK
        XlsxTestKit.Keep(path, "acceptance-test-procedure.docx");
#endif
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            Assert.Equal(3, main.ImageParts.Count());
            Assert.Single(main.HeaderParts);
            Assert.Single(main.HeaderParts.Single().ImageParts);
            Assert.Equal(2, main.FooterParts.Count());
            var tables = main.Document.Body!.Elements<Table>().ToList();
            Assert.True(tables.Count >= 12);
        }
    }

    [Fact]
    public void TablesTakeTheirStyleWidthsAndSpans()
    {
        var path = Save(Procedure(), "tables-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var tables = main.Document.Body!.Elements<Table>().ToList();
            var changes = tables[1];
            Assert.Equal("DivitTable", changes.GetFirstChild<TableProperties>()!.TableStyle!.Val!.Value);
            Assert.Equal(new[] { "850", "1417", "3969", "1417", "1417" }, changes.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Select(g => g.Width!.Value).ToArray());
            var header = changes.Elements<TableRow>().First();
            Assert.NotNull(header.TableRowProperties!.GetFirstChild<TableHeader>());
            Assert.Equal(JustificationValues.Center, changes.Elements<TableRow>().ElementAt(1).Elements<TableCell>().First().Elements<Paragraph>().First().ParagraphProperties!.Justification!.Val!.Value);

            // The record form: a section row over four columns, merged remarks, and the Grid look.
            var form = tables.Last();
            Assert.Equal("DivitGrid", form.GetFirstChild<TableProperties>()!.TableStyle!.Val!.Value);
            var spanned = form.Elements<TableRow>().ElementAt(2).Elements<TableCell>().Single();
            Assert.Equal(4, spanned.TableCellProperties!.GridSpan!.Val!.Value);
            Assert.Equal("BDD7EE", spanned.TableCellProperties.Shading!.Fill!.Value);
            Assert.Equal(2, form.Elements<TableRow>().ElementAt(7).Elements<TableCell>().Count());

            // Styles: the house table bands its rows and colours its header row.
            var styles = main.StyleDefinitionsPart!.Styles!;
            var house = styles.Elements<Style>().Single(s => s.StyleId == "DivitTable");
            var conditions = house.Elements<TableStyleProperties>().ToDictionary(p => p.Type!.Value);
            Assert.Equal("BDD7EE", conditions[TableStyleOverrideValues.FirstRow].TableStyleConditionalFormattingTableCellProperties!.Shading!.Fill!.Value);
            Assert.Equal("EEF5FB", conditions[TableStyleOverrideValues.Band1Horizontal].TableStyleConditionalFormattingTableCellProperties!.Shading!.Fill!.Value);
            Assert.DoesNotContain(styles.Elements<Style>().Single(s => s.StyleId == "DivitPlain").Elements<TableStyleProperties>(), p => true);
        }
    }

    [Fact]
    public void RowSpansContinueDownwards()
    {
        var doc = new DocxDocument();
        var table = doc.AddTable(3);
        var first = table.AddRow("Group", "a", "1");
        first[1].RowSpan = 3;
        table.AddRow(null, "b", "2");
        table.AddRow(null, "c", "3");
        table.AddRow("Next", "d", "4");
        var path = Save(doc, "rowspan-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var rows = document.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
            Assert.Equal(MergedCellValues.Restart, rows[0].Elements<TableCell>().First().TableCellProperties!.VerticalMerge!.Val!.Value);
            Assert.Null(rows[1].Elements<TableCell>().First().TableCellProperties!.VerticalMerge!.Val);
            Assert.NotNull(rows[2].Elements<TableCell>().First().TableCellProperties!.VerticalMerge);
            Assert.Null(rows[3].Elements<TableCell>().First().TableCellProperties!.VerticalMerge);
            Assert.All(rows, r => Assert.Equal(3, r.Elements<TableCell>().Count()));
        }
    }

    [Fact]
    public void ImagesLiveInThePartThatShowsThem()
    {
        var path = Save(Procedure(), "images-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var drawings = main.Document.Body!.Descendants<Drawing>().ToList();
            Assert.Equal(3, drawings.Count);
            var ids = main.Document.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
                .Concat(main.HeaderParts.SelectMany(h => h.Header.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()))
                .Select(p => p.Id!.Value).ToList();
            Assert.Equal(ids.Count, ids.Distinct().Count());
            Assert.Equal("SH2000 seen from the side", main.Document.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().ElementAt(1).Description!.Value);

            // The logo on the cover is 6 cm wide; its height follows the picture's proportions (3:1).
            var extent = drawings[0].Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().Single();
            Assert.Equal((2160000L, 720000L), (extent.Cx!.Value, extent.Cy!.Value));

            var blip = main.HeaderParts.Single().Header.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single();
            Assert.IsType<ImagePart>(main.HeaderParts.Single().GetPartById(blip.Embed!.Value!));
        }
    }

    [Fact]
    public void CaptionsAreNumberedAndListed()
    {
        var path = Save(Procedure(), "captions-check.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document.Body!;
            var sequences = body.Descendants<SimpleField>().Where(f => f.Instruction!.Value!.Contains("SEQ")).ToList();
            Assert.Equal(" SEQ Figure \\* ARABIC ", sequences.First(f => f.Instruction!.Value!.Contains("Figure")).Instruction!.Value);
            Assert.Equal(new[] { "1", "2" }, sequences.Where(f => f.Instruction!.Value!.Contains("Figure")).Select(f => f.InnerText).ToArray());
            Assert.Equal(12, sequences.Count(f => f.Instruction!.Value!.Contains("Table")));

            var captions = body.Elements<Paragraph>().Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Caption").Select(p => p.InnerText).ToList();
            Assert.Contains("Figure 2 – Return loss test set-up", captions);
            Assert.Contains("Table 1 – Document change record", captions);

            var lists = body.Descendants<FieldCode>().Select(f => f.Text).ToList();
            Assert.Contains(" TOC \\h \\z \\c \"Figure\" ", lists);
            Assert.Contains(" TOC \\h \\z \\c \"Table\" ", lists);
            var figures = body.Elements<Paragraph>().Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TableofFigures")
                .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToList();
            Assert.Equal("Figure 1 – SH2000 dual-ridge horn antenna", figures[0]);
            Assert.Equal(2 + 12, figures.Count);
        }
    }

    [Fact]
    public void TablesCheckTheirShape()
    {
        var doc = new DocxDocument();
        Assert.Throws<ArgumentOutOfRangeException>(() => doc.AddTable(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => doc.AddTable(64));
        var table = doc.AddTable(2);
        Assert.Throws<ArgumentException>(() => table.SetColumnWidths(1, 2, 3));
        Assert.Throws<ArgumentException>(() => table.AddRow("a", "b", "c"));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.AlignColumn(3, DocxAlignment.Right));
        Assert.Throws<ArgumentException>(() => doc.AddParagraph().AddImage(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public void TheProcedureIsReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        Procedure().Save(Path.Combine(folder, "a.docx"));
        Procedure().Save(Path.Combine(folder, "b.docx"));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.docx")), File.ReadAllBytes(Path.Combine(folder, "b.docx")));
    }
}
