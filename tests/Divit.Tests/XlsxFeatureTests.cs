namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Xunit;

/// <summary>
/// The smaller features one by one: every test saves a demo file, validates it and checks what Excel will read. The
/// demos land in the gallery for a look in Excel.
/// </summary>
public class XlsxFeatureTests
{
    /// <summary>
    /// Saves, validates, keeps the file in the gallery and opens it for checks.
    /// </summary>
    private static SpreadsheetDocument SaveAndOpen(XlsxWorkbook book, string name)
    {
        book.Created ??= XlsxTestKit.FixedTime;
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        XlsxTestKit.Keep(path, name);
        return SpreadsheetDocument.Open(path, false);
    }

    private static Worksheet SheetOf(SpreadsheetDocument document, string name)
    {
        var part = document.WorkbookPart!;
        var sheet = part.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == name);
        return ((WorksheetPart)part.GetPartById(sheet.Id!)).Worksheet;
    }

    private static Cell CellOf(Worksheet sheet, string reference) =>
        sheet.Descendants<Cell>().Single(c => c.CellReference == reference);

    [Fact]
    public void DatesAndBooleans()
    {
        var book = new XlsxWorkbook { Title = "Dates and booleans" };
        var sheet = book.AddSheet("Dates");
        sheet.Set(1, 1, "What");
        sheet.Set(1, 2, "Value");
        sheet.Set(1, 3, "Passed");
        sheet.Format(1, 1, 1, 3, (ref XlsxStyle s) => s.Bold = true);
        sheet.Set(2, 1, "Date");
        sheet.Set(2, 2, new DateTime(2026, 11, 7));
        sheet.Set(2, 3, true);
        sheet.Set(3, 1, "Date and time");
        sheet.Set(3, 2, new DateTime(2026, 11, 7, 18, 54, 2));
        sheet.Set(3, 3, false);
        sheet.Set(4, 1, "Own format");
        sheet.Format(4, 2, (ref XlsxStyle s) => s.NumberFormat = "d mmmm yyyy");
        sheet.Set(4, 2, new DateTime(1923, 10, 29));
        sheet.Set(4, 3, true);
        sheet.ColumnWidth(1, 16);
        sheet.ColumnWidth(2, 22);

        using (var document = SaveAndOpen(book, "dates-booleans.xlsx"))
        {
            var ws = SheetOf(document, "Dates");
            var date = CellOf(ws, "B2");
            Assert.Equal(new DateTime(2026, 11, 7).ToOADate(), double.Parse(date.CellValue!.Text, CultureInfo.InvariantCulture));
            Assert.Equal(CellValues.Boolean, CellOf(ws, "C2").DataType!.Value);
            Assert.Equal("1", CellOf(ws, "C2").CellValue!.Text);
            Assert.Equal("0", CellOf(ws, "C3").CellValue!.Text);

            var formats = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet.NumberingFormats!.Elements<NumberingFormat>()
                .Select(f => f.FormatCode!.Value).ToList();
            Assert.Contains("yyyy-mm-dd", formats);
            Assert.Contains("yyyy-mm-dd hh:mm:ss", formats);
            Assert.Contains("d mmmm yyyy", formats);
        }
    }

    [Theory]
    [InlineData(1900, 1, 1, 1)]
    [InlineData(1900, 2, 28, 59)]
    [InlineData(1900, 3, 1, 61)]
    [InlineData(2026, 1, 2, 46024)]
    public void DateNumbersFollowExcel(int year, int month, int day, double expected)
    {
        Assert.Equal(expected, XlsxSheet.DateNumber(new DateTime(year, month, day)));
    }

    [Fact]
    public void DatesBefore1900AreRejected()
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.Set(1, 1, new DateTime(1899, 12, 31)));
    }

    [Fact]
    public void Comments()
    {
        var book = new XlsxWorkbook { Author = "Divit tests" };
        var sheet = book.AddSheet("Notes");
        sheet.Set(1, 1, "Hover over the cells with a red corner.");
        sheet.Set(3, 2, 42.0);
        sheet.Comment(3, 2, "The answer.\nSecond line.");
        sheet.Set(5, 4, "Measured");
        sheet.Comment(5, 4, "Checked by the lab.", "Lab");
        sheet.Comment(1, 1, "A note on the first cell.");
        var other = book.AddSheet("More notes");
        other.Comment(2, 2, "A note on the second sheet.");

        using (var document = SaveAndOpen(book, "comments.xlsx"))
        {
            var parts = document.WorkbookPart!.WorksheetParts.ToList();
            var notes = parts.Select(p => p.WorksheetCommentsPart).Where(p => p != null).ToList();
            Assert.Equal(2, notes.Count);
            Assert.All(parts, p => Assert.Single(p.VmlDrawingParts));

            var first = SheetOf(document, "Notes").WorksheetPart!.WorksheetCommentsPart!.Comments;
            Assert.Equal(new[] { "A1", "B3", "D5" }, first.CommentList!.Elements<Comment>().Select(c => c.Reference!.Value).ToArray());
            Assert.Equal(new[] { "Divit tests", "Lab" }, first.Authors!.Elements<Author>().Select(a => a.Text).ToArray());
        }
    }

    [Fact]
    public void HiddenAndGroupedRowsAndColumns()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("Outline");
        sheet.Set(1, 1, "Region");
        sheet.Set(1, 2, "Q1");
        sheet.Set(1, 3, "Q2");
        sheet.Set(1, 4, "Total");
        var row = 2;
        foreach (var region in new[] { "North", "South" })
        {
            var first = row;
            for (var city = 1; city <= 3; city++, row++)
            {
                sheet.Set(row, 1, region + " city " + city);
                sheet.Set(row, 2, city * 10.0);
                sheet.Set(row, 3, city * 12.0);
            }

            sheet.GroupRows(first, row - 1, collapsed: region == "South");
            sheet.Set(row, 1, region + " total");
            sheet.Format(row, 1, row, 4, (ref XlsxStyle s) => s.Bold = true);
            row++;
        }

        sheet.GroupColumns(2, 3);
        sheet.Set(row + 1, 1, "Row below is hidden");
        sheet.Set(row + 2, 1, "hidden");
        sheet.HideRow(row + 2);

        using (var document = SaveAndOpen(book, "outline.xlsx"))
        {
            var ws = SheetOf(document, "Outline");
            var rows = ws.Descendants<Row>().ToDictionary(r => (int)r.RowIndex!.Value);
            Assert.Equal(1, rows[2].OutlineLevel!.Value);
            Assert.Null(rows[2].Hidden);
            Assert.True(rows[6].Hidden!.Value);
            Assert.True(rows[9].Collapsed!.Value);
            Assert.True(rows[row + 2].Hidden!.Value);
            Assert.Equal(1, ws.GetFirstChild<SheetFormatProperties>()!.OutlineLevelRow!.Value);
            Assert.Equal(1, ws.GetFirstChild<SheetFormatProperties>()!.OutlineLevelColumn!.Value);
            Assert.Equal(2, ws.Descendants<Column>().Count(c => c.OutlineLevel?.Value == 1));
        }
    }

    [Fact]
    public void SheetVisibilityOrderAndActiveTab()
    {
        var book = new XlsxWorkbook();
        var lookup = book.AddSheet("Lookup");
        lookup.Visibility = XlsxSheetVisibility.VeryHidden;
        lookup.Set(1, 1, "Only a program can show this sheet.");
        var draft = book.AddSheet("Draft");
        draft.Visibility = XlsxSheetVisibility.Hidden;
        draft.Set(1, 1, "Unhide shows this sheet.");
        var report = book.AddSheet("Report");
        report.Set(1, 1, "This sheet opens first.");
        var intro = book.AddSheet("Intro");
        intro.Set(1, 1, "Moved to the front of the visible tabs.");
        book.MoveSheet(intro, 2);
        book.ActiveSheet = report;

        using (var document = SaveAndOpen(book, "sheets.xlsx"))
        {
            var workbook = document.WorkbookPart!.Workbook;
            var sheets = workbook.Sheets!.Elements<Sheet>().ToList();
            Assert.Equal(new[] { "Lookup", "Draft", "Intro", "Report" }, sheets.Select(s => s.Name!.Value).ToArray());
            Assert.Equal(SheetStateValues.VeryHidden, sheets[0].State!.Value);
            Assert.Equal(SheetStateValues.Hidden, sheets[1].State!.Value);
            Assert.Null(sheets[3].State);
            var view = workbook.BookViews!.GetFirstChild<WorkbookView>()!;
            Assert.Equal(3u, view.ActiveTab!.Value);
            Assert.Equal(2u, view.FirstSheet!.Value);
            Assert.True(SheetOf(document, "Report").Descendants<SheetView>().Single().TabSelected!.Value);
        }
    }

    [Fact]
    public void HiddenSheetsCannotOpenFirst()
    {
        var book = new XlsxWorkbook();
        var only = book.AddSheet("Only");
        only.Visibility = XlsxSheetVisibility.Hidden;
        var path = Path.Combine(XlsxTestKit.NewFolder(), "hidden.xlsx");
        Assert.Throws<InvalidOperationException>(() => book.Save(path));

        var shown = book.AddSheet("Shown");
        book.ActiveSheet = only;
        Assert.Throws<InvalidOperationException>(() => book.Save(path));
        book.ActiveSheet = shown;
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
    }

    [Fact]
    public void DefinedNames()
    {
        var book = new XlsxWorkbook();
        var data = book.AddSheet("Sales data");
        data.Set(1, 1, "Month");
        data.Set(1, 2, "Sales");
        for (var i = 1; i <= 12; i++)
        {
            data.Set(i + 1, 1, CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(i));
            data.Set(i + 1, 2, 100.0 + (i * 7));
        }

        book.DefineName("Sales", data, 2, 2, 13, 2);
        book.DefineName("FirstMonth", data, 2, 1, 2, 1);
        var summary = book.AddSheet("Summary");
        summary.Set(1, 1, "Pick Sales or FirstMonth in the Name Box to jump to the range.");
        book.DefineName("Sales", data, 2, 2, 7, 2, scope: summary);

        using (var document = SaveAndOpen(book, "names.xlsx"))
        {
            var names = document.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
            Assert.Equal("'Sales data'!$B$2:$B$13", names.Single(n => n.Name == "Sales" && n.LocalSheetId == null).Text);
            Assert.Equal("'Sales data'!$A$2", names.Single(n => n.Name == "FirstMonth").Text);
            Assert.Equal(1u, names.Single(n => n.LocalSheetId != null).LocalSheetId!.Value);
        }
    }

    [Theory]
    [InlineData("Sales", true)]
    [InlineData("_total", true)]
    [InlineData("Tax.Rate", true)]
    [InlineData("Ölçüm", true)]
    [InlineData("XFE1", true)]
    [InlineData("", false)]
    [InlineData("1st", false)]
    [InlineData("Two words", false)]
    [InlineData("A1", false)]
    [InlineData("xfd1048576", false)]
    [InlineData("R", false)]
    [InlineData("c", false)]
    [InlineData("R1C1", false)]
    [InlineData("R12", false)]
    [InlineData("_xlnm.Print_Area", false)]
    public void DefinedNamesAreChecked(string name, bool usable)
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        if (usable)
        {
            book.DefineName(name, sheet, 1, 1, 1, 1);
        }
        else
        {
            Assert.Throws<ArgumentException>(() => book.DefineName(name, sheet, 1, 1, 1, 1));
        }
    }

    [Fact]
    public void DuplicateDefinedNamesIgnoreCase()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        book.DefineName("Total", sheet, 1, 1, 1, 1);
        Assert.Throws<ArgumentException>(() => book.DefineName("TOTAL", sheet, 2, 1, 2, 1));
        Assert.Throws<ArgumentException>(() => book.DefineName("Other", new XlsxWorkbook().AddSheet("Two"), 1, 1, 1, 1));
    }

    [Fact]
    public void PageSetup()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("Print");
        for (var r = 1; r <= 60; r++)
        {
            for (var c = 1; c <= 6; c++)
            {
                sheet.Set(r, c, r * c);
            }
        }

        sheet.PaperSize = XlsxPaperSize.A5;
        sheet.Landscape = true;
        sheet.Margins = XlsxMargins.FromCentimeters(1, 1, 2, 2, 0.8, 0.8);
        sheet.PrintCenteredHorizontally = true;
        sheet.PrintGridLines = true;
        sheet.Footer = "&CPage &P / &N";
        var letter = book.AddSheet("Letter");
        letter.Set(1, 1, "Letter paper, Excel's wide margins.");
        letter.PaperSize = XlsxPaperSize.Letter;
        letter.Margins = XlsxMargins.Wide;

        using (var document = SaveAndOpen(book, "print-setup.xlsx"))
        {
            var ws = SheetOf(document, "Print");
            Assert.Equal(11u, ws.GetFirstChild<PageSetup>()!.PaperSize!.Value);
            Assert.Equal(1 / 2.54, ws.GetFirstChild<PageMargins>()!.Left!.Value, 4);
            Assert.True(ws.GetFirstChild<PrintOptions>()!.HorizontalCentered!.Value);
            Assert.True(ws.GetFirstChild<PrintOptions>()!.GridLines!.Value);
            Assert.Equal(1u, SheetOf(document, "Letter").GetFirstChild<PageSetup>()!.PaperSize!.Value);
            Assert.Null(SheetOf(document, "Letter").GetFirstChild<PrintOptions>());
        }
    }

    [Fact]
    public void FontsAndAlignment()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("Look");
        void Row(int row, string label, XlsxStyleEdit edit)
        {
            sheet.Set(row, 1, label);
            sheet.Set(row, 2, "Sample text");
            sheet.Format(row, 2, edit);
        }

        Row(1, "Strikethrough", (ref XlsxStyle s) => s.Strikethrough = true);
        Row(2, "Double underline", (ref XlsxStyle s) => s.DoubleUnderline = true);
        Row(3, "Superscript", (ref XlsxStyle s) => s.Script = XlsxScript.Superscript);
        Row(4, "Subscript", (ref XlsxStyle s) => s.Script = XlsxScript.Subscript);
        Row(5, "Indent 2", (ref XlsxStyle s) => s.Indent = 2);
        Row(6, "Right, indent 1", (ref XlsxStyle s) => { s.Horizontal = "right"; s.Indent = 1; });
        Row(7, "Rotated 45°", (ref XlsxStyle s) => s.Rotation = 45);
        Row(8, "Rotated -45°", (ref XlsxStyle s) => s.Rotation = -45);
        Row(9, "Vertical", (ref XlsxStyle s) => s.Rotation = XlsxStyle.VerticalText);
        Row(10, "Shrink to fit", (ref XlsxStyle s) => s.ShrinkToFit = true);
        sheet.Set(10, 2, "A long text that shrinks to fit the cell");
        sheet.ColumnWidth(1, 18);
        sheet.ColumnWidth(2, 16);
        for (var r = 7; r <= 9; r++)
        {
            sheet.RowHeight(r, 70);
        }

        using (var document = SaveAndOpen(book, "fonts-alignment.xlsx"))
        {
            var styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            var fonts = styles.Fonts!.Elements<Font>().ToList();
            Assert.Contains(fonts, f => f.Strike != null);
            Assert.Contains(fonts, f => f.Underline?.Val?.Value == UnderlineValues.Double);
            Assert.Contains(fonts, f => f.VerticalTextAlignment?.Val?.Value == VerticalAlignmentRunValues.Superscript);
            Assert.Contains(fonts, f => f.VerticalTextAlignment?.Val?.Value == VerticalAlignmentRunValues.Subscript);

            var alignments = styles.CellFormats!.Elements<CellFormat>().Select(f => f.Alignment).Where(a => a != null).ToList();
            Assert.Contains(alignments, a => a!.Indent?.Value == 2 && a.Horizontal!.Value == HorizontalAlignmentValues.Left);
            Assert.Contains(alignments, a => a!.TextRotation?.Value == 45);
            Assert.Contains(alignments, a => a!.TextRotation?.Value == 135);
            Assert.Contains(alignments, a => a!.TextRotation?.Value == 255);
            Assert.Contains(alignments, a => a!.ShrinkToFit?.Value == true);
        }
    }

    [Fact]
    public void ExternalLinks()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("Links");
        sheet.LinkUrl(1, 1, "https://github.com/karci199", "Web page");
        sheet.LinkUrl(2, 1, "mailto:someone@example.com", "E-mail");
        sheet.LinkUrl(3, 1, "full.xlsx", "File next to this one");
        sheet.LinkUrl(4, 1, "https://example.com/a?b=1&c=2");
        sheet.Set(5, 1, "Back to the top");
        sheet.Link(5, 1, "Links");
        sheet.ColumnWidth(1, 32);

        using (var document = SaveAndOpen(book, "links.xlsx"))
        {
            var part = SheetOf(document, "Links").WorksheetPart!;
            var links = part.HyperlinkRelationships.ToList();
            Assert.Equal(4, links.Count);
            Assert.All(links, l => Assert.True(l.IsExternal));
            Assert.Contains(links, l => l.Uri.OriginalString == "mailto:someone@example.com");
            Assert.Contains(links, l => l.Uri.OriginalString == "full.xlsx");
            Assert.Equal(5, part.Worksheet.Descendants<Hyperlink>().Count());
            // Without text the cell shows the address itself.
            Assert.Contains("https://example.com/a?b=1&c=2", document.WorkbookPart!.SharedStringTablePart!.SharedStringTable
                .Elements<SharedStringItem>().Select(s => s.InnerText));
        }

        Assert.Throws<ArgumentException>(() => sheet.LinkUrl(9, 1, string.Empty));
    }
}
