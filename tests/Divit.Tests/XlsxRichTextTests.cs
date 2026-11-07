namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;
using Xunit;

/// <summary>
/// Rich text cells: pieces with their own fonts in the shared string table.
/// </summary>
public class XlsxRichTextTests
{
    private static readonly XlsxStyle Bold = new XlsxStyle { Bold = true };
    private static readonly XlsxStyle RedBold = new XlsxStyle { Bold = true, FontColor = "#C00000" };

    private static XlsxWorkbook Samples()
    {
        var book = new XlsxWorkbook { Title = "Rich text", Created = XlsxTestKit.FixedTime };
        var sheet = book.AddSheet("Rich text");
        sheet.SetRichText(1, 1, "Status: ", new XlsxRun("late", RedBold), " (3 days)");
        sheet.SetRichText(2, 1, "Area: 12 m", new XlsxRun("2", new XlsxStyle { Script = XlsxScript.Superscript }));
        sheet.SetRichText(3, 1, "Water: H", new XlsxRun("2", new XlsxStyle { Script = XlsxScript.Subscript }), "O");
        sheet.SetRichText(4, 1, new XlsxRun("Big ", new XlsxStyle { FontSize = 16 }), new XlsxRun("small", new XlsxStyle { FontSize = 8 }));
        sheet.SetRichText(5, 1, new XlsxRun("Consolas", new XlsxStyle { FontName = "Consolas" }), " next to the default font");
        sheet.SetRichText(6, 1, new XlsxRun("old price", new XlsxStyle { Strikethrough = true }), " ", new XlsxRun("new price", new XlsxStyle { DoubleUnderline = true }));
        sheet.SetRichText(7, 1, new XlsxRun("First line", Bold), "\nsecond line");
        sheet.Format(7, 1, (ref XlsxStyle s) => s.Wrap = true);
        sheet.RowHeight(7, 30);
        sheet.SetRichText(8, 1, "Status: ", new XlsxRun("late", RedBold), " (3 days)");
        sheet.SetRichText(9, 1, "  spaced  ", new XlsxRun("end ", Bold));
        sheet.ColumnWidth(1, 40);
        return book;
    }

    private static SharedStringItem[] Strings(SpreadsheetDocument document) =>
        document.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ToArray();

    private static SharedStringItem ItemOf(SpreadsheetDocument document, string reference)
    {
        var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single(c => c.CellReference == reference);
        Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
        return Strings(document)[int.Parse(cell.CellValue!.Text)];
    }

    [Fact]
    public void PiecesKeepTheirFonts()
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), "rich-text.xlsx");
        Samples().Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        XlsxTestKit.Keep(path, "rich-text.xlsx");

        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var status = ItemOf(document, "A1").Elements<Run>().ToArray();
            Assert.Equal(new[] { "Status: ", "late", " (3 days)" }, status.Select(r => r.Text!.Text).ToArray());

            // A piece in the cell's font has no properties; a piece with a look has complete ones.
            Assert.Null(status[0].RunProperties);
            var late = status[1].RunProperties!;
            Assert.NotNull(late.GetFirstChild<Bold>());
            Assert.Equal("FFC00000", late.GetFirstChild<Color>()!.Rgb!.Value);
            Assert.Equal(10, late.GetFirstChild<FontSize>()!.Val!.Value);
            Assert.Equal("Segoe UI", late.GetFirstChild<RunFont>()!.Val!.Value);
            Assert.Null(late.GetFirstChild<Italic>());

            // Leading and trailing spaces survive.
            Assert.Equal(DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve, status[0].Text!.Space!.Value);

            var square = ItemOf(document, "A2").Elements<Run>().Last().RunProperties!;
            Assert.Equal(VerticalAlignmentRunValues.Superscript, square.GetFirstChild<VerticalTextAlignment>()!.Val!.Value);
            var water = ItemOf(document, "A3").Elements<Run>().ElementAt(1).RunProperties!;
            Assert.Equal(VerticalAlignmentRunValues.Subscript, water.GetFirstChild<VerticalTextAlignment>()!.Val!.Value);

            var sizes = ItemOf(document, "A4").Elements<Run>().Select(r => r.RunProperties!.GetFirstChild<FontSize>()!.Val!.Value).ToArray();
            Assert.Equal(new[] { 16.0, 8.0 }, sizes);
            Assert.Equal("Consolas", ItemOf(document, "A5").Elements<Run>().First().RunProperties!.GetFirstChild<RunFont>()!.Val!.Value);

            var prices = ItemOf(document, "A6").Elements<Run>().ToArray();
            Assert.NotNull(prices[0].RunProperties!.GetFirstChild<Strike>());
            Assert.Equal(UnderlineValues.Double, prices[2].RunProperties!.GetFirstChild<Underline>()!.Val!.Value);

            Assert.Contains("\n", ItemOf(document, "A7").InnerText);
            Assert.Equal("  spaced  end ", ItemOf(document, "A9").InnerText);
        }
    }

    [Fact]
    public void EqualRichTextsAreStoredOnce()
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), "rich-once.xlsx");
        Samples().Save(path);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
            Assert.Equal(cells["A1"].CellValue!.Text, cells["A8"].CellValue!.Text);
            var table = document.WorkbookPart.SharedStringTablePart!.SharedStringTable;
            Assert.Equal(8u, table.UniqueCount!.Value);
            Assert.Equal(9u, table.Count!.Value);
        }
    }

    [Fact]
    public void SameTextWithOtherFontsIsAnotherEntryAndPlainTextStaysPlain()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.Set(1, 1, "ab");
        sheet.SetRichText(2, 1, "a", new XlsxRun("b", Bold));
        sheet.SetRichText(3, 1, "a", new XlsxRun("b", RedBold));
        sheet.SetRichText(4, 1, "ab");
        sheet.SetRichText(5, 1, string.Empty, new XlsxRun(string.Empty, Bold));
        sheet.SetRichText(6, 1, new XlsxRun("x", Bold), string.Empty);
        var path = Path.Combine(XlsxTestKit.NewFolder(), "rich-entries.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var items = Strings(document);
            Assert.Equal(4, items.Length);
            Assert.Equal("ab", items[0].Text!.Text);
            Assert.Empty(items[0].Elements<Run>());
            Assert.Equal(2, items[1].Elements<Run>().Count());
            Assert.Single(items[3].Elements<Run>());

            // "ab" as one plain piece reuses the plain entry; all-empty pieces leave the cell empty.
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);
            Assert.Equal("0", cells["A4"].CellValue!.Text);
            Assert.False(cells.ContainsKey("A5"));
        }
    }

    [Fact]
    public void RichTextReplacesAFormula()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.SetFormula(1, 1, "1+1");
        sheet.SetRichText(1, 1, new XlsxRun("text", Bold));
        var path = Path.Combine(XlsxTestKit.NewFolder(), "rich-formula.xlsx");
        book.Save(path);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single();
            Assert.Null(cell.CellFormula);
            Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
        }
    }

    [Fact]
    public void RunsCompareByTextAndLook()
    {
        Assert.Equal(new XlsxRun("a", Bold), new XlsxRun("a", new XlsxStyle { Bold = true }));
        Assert.NotEqual(new XlsxRun("a", Bold), new XlsxRun("a"));
        Assert.NotEqual(new XlsxRun("a"), new XlsxRun("b"));
        XlsxRun implicitRun = "c";
        Assert.Null(implicitRun.Look);
    }
}
