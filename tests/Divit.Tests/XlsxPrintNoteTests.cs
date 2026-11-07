namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Xunit;

/// <summary>
/// Manual page breaks, first and even page headers, and notes with fonts and their own box size.
/// </summary>
public class XlsxPrintNoteTests
{
    private static XlsxWorkbook Report()
    {
        var book = new XlsxWorkbook { Title = "Page breaks, headers and notes", Created = XlsxTestKit.FixedTime };
        var sheet = book.AddSheet("Report");
        for (var r = 1; r <= 90; r++)
        {
            sheet.Set(r, 1, "Line " + r);
            for (var c = 2; c <= 12; c++)
            {
                sheet.Set(r, c, r * c);
            }
        }

        sheet.BreakAfterRow(30);
        sheet.BreakAfterRow(60);
        sheet.BreakAfterRow(30);
        sheet.BreakAfterColumn(6);
        sheet.Header = "&LDivit&RPage &P";
        sheet.Footer = "&C&A";
        sheet.FirstPageHeader = "&C&\"-,Bold\"Annual report";
        sheet.FirstPageFooter = string.Empty;
        sheet.EvenPageHeader = "&LPage &P&RDivit";
        sheet.EvenPageFooter = "&C&A";

        sheet.Set(1, 14, "Hover over N1 and N3.");
        sheet.RichComment(1, 14, "Lab", new XlsxRun("Checked\n", new XlsxStyle { Bold = true }), "by the lab on ",
            new XlsxRun("7 November", new XlsxStyle { FontColor = "#C00000" }));
        var wide = sheet.Comment(3, 14, "A wide note: four columns and two rows.");
        wide.Columns = 4;
        wide.Rows = 2;

        var fitted = book.AddSheet("Fitted");
        fitted.Set(1, 1, "Fitted to one page wide: no manual breaks.");
        return book;
    }

    private static string Save(out string path)
    {
        path = Path.Combine(XlsxTestKit.NewFolder(), "print-notes.xlsx");
        Report().Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        return path;
    }

    private static Worksheet SheetOf(SpreadsheetDocument document, string name)
    {
        var part = document.WorkbookPart!;
        var id = part.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == name).Id!.Value!;
        return ((WorksheetPart)part.GetPartById(id)).Worksheet;
    }

    [Fact]
    public void PageBreaksTurnFittingOff()
    {
        Save(out var path);
        XlsxTestKit.Keep(path, "print-notes.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sheet = SheetOf(document, "Report");
            var rows = sheet.GetFirstChild<RowBreaks>()!;
            Assert.Equal(2u, rows.Count!.Value);
            Assert.Equal(2u, rows.ManualBreakCount!.Value);
            var breaks = rows.Elements<Break>().ToList();
            Assert.Equal(new uint[] { 30, 60 }, breaks.Select(b => b.Id!.Value).ToArray());
            Assert.All(breaks, b => Assert.Equal(16383u, b.Max!.Value));
            Assert.All(breaks, b => Assert.True(b.ManualPageBreak!.Value));
            var columns = sheet.GetFirstChild<ColumnBreaks>()!.Elements<Break>().Single();
            Assert.Equal((6u, 1048575u), (columns.Id!.Value, columns.Max!.Value));

            Assert.False(sheet.SheetProperties!.PageSetupProperties!.FitToPage!.Value);
            Assert.True(SheetOf(document, "Fitted").SheetProperties!.PageSetupProperties!.FitToPage!.Value);
            Assert.Null(SheetOf(document, "Fitted").GetFirstChild<RowBreaks>());
        }
    }

    [Fact]
    public void FirstAndEvenPagesHaveTheirOwnHeaders()
    {
        Save(out var path);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var headers = SheetOf(document, "Report").GetFirstChild<HeaderFooter>()!;
            Assert.True(headers.DifferentFirst!.Value);
            Assert.True(headers.DifferentOddEven!.Value);
            Assert.Equal("&LDivit&RPage &P", headers.OddHeader!.Text);
            Assert.Equal("&LPage &P&RDivit", headers.EvenHeader!.Text);
            Assert.Equal("&C&\"-,Bold\"Annual report", headers.FirstHeader!.Text);
            Assert.Equal(string.Empty, headers.FirstFooter!.Text);
            Assert.Null(SheetOf(document, "Fitted").GetFirstChild<HeaderFooter>());
        }
    }

    [Fact]
    public void OnlyTheHeadersGivenAreWritten()
    {
        var book = new XlsxWorkbook();
        book.AddSheet("One").FirstPageFooter = "&CFirst only";
        var path = Path.Combine(XlsxTestKit.NewFolder(), "first-only.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var headers = document.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<HeaderFooter>()!;
            Assert.True(headers.DifferentFirst!.Value);
            Assert.Null(headers.DifferentOddEven);
            Assert.Null(headers.OddHeader);
            Assert.Equal("&CFirst only", headers.FirstFooter!.Text);
        }
    }

    [Fact]
    public void NotesKeepTheirFontsAndSize()
    {
        Save(out var path);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var comments = SheetOf(document, "Report").WorksheetPart!.WorksheetCommentsPart!.Comments;
            var rich = comments.CommentList!.Elements<Comment>().First();
            var runs = rich.CommentText!.Elements<Run>().ToList();
            Assert.Equal(new[] { "Checked\n", "by the lab on ", "7 November" }, runs.Select(r => r.Text!.Text).ToArray());
            Assert.NotNull(runs[0].RunProperties!.GetFirstChild<Bold>());
            Assert.Null(runs[1].RunProperties);
            Assert.Equal("FFC00000", runs[2].RunProperties!.GetFirstChild<Color>()!.Rgb!.Value);
            Assert.Equal("Lab", comments.Authors!.Elements<Author>().ElementAt((int)rich.AuthorId!.Value).Text);
        }

        // The VML box: two columns by four rows by default, four by two for the wide note.
        using (var zip = new ZipArchive(File.OpenRead(path), ZipArchiveMode.Read))
        using (var reader = new StreamReader(zip.Entries.Single(e => e.FullName.EndsWith(".vml", System.StringComparison.Ordinal)).Open()))
        {
            var anchors = System.Text.RegularExpressions.Regex.Matches(reader.ReadToEnd(), "<x:Anchor>([^<]*)</x:Anchor>")
                .Cast<System.Text.RegularExpressions.Match>().Select(m => m.Groups[1].Value).ToArray();
            Assert.Equal(new[] { "14, 15, 0, 10, 16, 15, 4, 4", "14, 15, 1, 10, 18, 15, 3, 4" }, anchors);
        }
    }

    [Fact]
    public void NoteSizesAreClamped()
    {
        var note = new XlsxWorkbook().AddSheet("One").Comment(1, 1, "x");
        note.Columns = 0;
        note.Rows = 500;
        Assert.Equal((1, 100), (note.Columns, note.Rows));
        Assert.Equal(2, new XlsxWorkbook().AddSheet("Two").Comment(1, 1, "y").Columns);
    }
}
