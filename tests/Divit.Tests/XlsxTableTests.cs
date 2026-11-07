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
/// Excel tables (ListObjects): the table part, its columns and totals row, the cells it fills, and every rule Excel
/// sets for tables.
/// </summary>
public class XlsxTableTests
{
    private static readonly string[] Regions = { "North", "South", "East", "West" };

    private static string Save(XlsxWorkbook book, string name, bool keep = false)
    {
        book.Created ??= XlsxTestKit.FixedTime;
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    /// <summary>
    /// A sales table with a totals row, a small price table without filter buttons, and a table whose headers need
    /// fixing; a formula outside uses the sales table by name.
    /// </summary>
    private static XlsxWorkbook Sales()
    {
        var book = new XlsxWorkbook { Title = "Tables" };
        var sheet = book.AddSheet("Sales");
        var headers = new[] { "Date", "Region", "Product", "Quantity", "Amount" };
        for (var c = 0; c < headers.Length; c++)
        {
            sheet.Set(1, c + 1, headers[c]);
        }

        for (var r = 0; r < 12; r++)
        {
            var row = r + 2;
            sheet.Set(row, 1, new DateTime(2026, r + 1, 15));
            sheet.Set(row, 2, Regions[r % 4]);
            sheet.Set(row, 3, r % 3 == 0 ? "Pen" : "Paper");
            sheet.Set(row, 4, 10 + r);
            sheet.Set(row, 5, (10 + r) * 2.5);
        }

        sheet.Format(2, 5, 14, 5, (ref XlsxStyle s) => s.NumberFormat = "#,##0.00");
        var sales = sheet.AddTable(1, 1, 13, 5, "Sales");
        sales.Style = "TableStyleMedium9";
        sales.TotalLabel(1, "Total");
        sales.Total(2, XlsxTotal.Count);
        sales.Total(4, XlsxTotal.Average);
        sales.Total(5, XlsxTotal.Sum);
        sheet.Set(16, 4, "Sum by name");
        sheet.SetFormula(16, 5, "SUM(Sales[Amount])", 465.0);
        for (var c = 1; c <= 5; c++)
        {
            sheet.ColumnWidth(c, 13);
        }

        var prices = sheet.AddTable(1, 8, 3, 9, "Prices");
        sheet.Set(1, 8, "Product");
        sheet.Set(1, 9, "Price [USD]");
        sheet.Set(2, 8, "Pen");
        sheet.Set(2, 9, 1.5);
        sheet.Set(3, 8, "Paper");
        sheet.Set(3, 9, 4.25);
        prices.Style = "TableStyleLight1";
        prices.FilterButtons = false;
        prices.RowStripes = false;
        prices.ColumnStripes = true;
        prices.FirstColumn = true;
        prices.Total(2, XlsxTotal.Max);

        var raw = book.AddSheet("Raw");
        raw.Set(1, 1, "Name");
        raw.Set(1, 3, "Name");
        raw.Set(1, 4, 2026.0);
        for (var r = 2; r <= 4; r++)
        {
            for (var c = 1; c <= 4; c++)
            {
                raw.Set(r, c, r * c);
            }
        }

        var plain = raw.AddTable(1, 1, 4, 4, "Raw_data");
        plain.Style = null;
        plain.LastColumn = true;
        return book;
    }

    private static TableDefinitionPart TableOf(SpreadsheetDocument document, string name) =>
        document.WorkbookPart!.WorksheetParts.SelectMany(p => p.TableDefinitionParts).Single(t => t.Table.Name == name);

    private static Cell CellOf(SpreadsheetDocument document, string sheet, string reference)
    {
        var part = document.WorkbookPart!;
        var id = part.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == sheet).Id!.Value!;
        return ((WorksheetPart)part.GetPartById(id)).Worksheet.Descendants<Cell>().Single(c => c.CellReference == reference);
    }

    private static string TextOf(SpreadsheetDocument document, Cell cell) =>
        cell.DataType?.Value == CellValues.SharedString
            ? document.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue!.Text)).InnerText
            : cell.CellValue!.Text;

    [Fact]
    public void TablesWithTotalsAndStyles()
    {
        var path = Save(Sales(), "tables.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sales = TableOf(document, "Sales").Table;
            Assert.Equal("Sales", sales.DisplayName!.Value);
            Assert.Equal("A1:E14", sales.Reference!.Value);
            Assert.Equal(1u, sales.TotalsRowCount!.Value);
            Assert.Equal("A1:E13", sales.AutoFilter!.Reference!.Value);
            var columns = sales.TableColumns!.Elements<TableColumn>().ToList();
            Assert.Equal(new[] { "Date", "Region", "Product", "Quantity", "Amount" }, columns.Select(c => c.Name!.Value).ToArray());
            Assert.Equal(Enumerable.Range(1, 5).Select(i => (uint)i), columns.Select(c => c.Id!.Value));
            Assert.Equal("Total", columns[0].TotalsRowLabel!.Value);
            Assert.Equal(TotalsRowFunctionValues.Count, columns[1].TotalsRowFunction!.Value);
            Assert.Null(columns[2].TotalsRowFunction);
            Assert.Equal(TotalsRowFunctionValues.Average, columns[3].TotalsRowFunction!.Value);
            Assert.Equal(TotalsRowFunctionValues.Sum, columns[4].TotalsRowFunction!.Value);
            var style = sales.TableStyleInfo!;
            Assert.Equal("TableStyleMedium9", style.Name!.Value);
            Assert.True(style.ShowRowStripes!.Value);
            Assert.False(style.ShowColumnStripes!.Value);

            // The totals row: a text label and SUBTOTAL formulas with their results.
            Assert.Equal("Total", TextOf(document, CellOf(document, "Sales", "A14")));
            var sum = CellOf(document, "Sales", "E14");
            Assert.Equal("SUBTOTAL(109,Sales[Amount])", sum.CellFormula!.Text);
            Assert.Equal(465.0, double.Parse(sum.CellValue!.Text, CultureInfo.InvariantCulture));
            Assert.Equal("SUBTOTAL(103,Sales[Region])", CellOf(document, "Sales", "B14").CellFormula!.Text);
            Assert.Equal("12", CellOf(document, "Sales", "B14").CellValue!.Text);
            Assert.Equal(15.5, double.Parse(CellOf(document, "Sales", "D14").CellValue!.Text, CultureInfo.InvariantCulture));
            Assert.Equal("SUM(Sales[Amount])", CellOf(document, "Sales", "E16").CellFormula!.Text);

            // Special characters in a column name are escaped in the structured reference.
            var prices = TableOf(document, "Prices").Table;
            Assert.Null(prices.AutoFilter);
            Assert.Equal("H1:I4", prices.Reference!.Value);
            Assert.Equal("SUBTOTAL(104,Prices[Price '[USD']])", CellOf(document, "Sales", "I4").CellFormula!.Text);
            Assert.Equal("4.25", CellOf(document, "Sales", "I4").CellValue!.Text);
            Assert.True(prices.TableStyleInfo!.ShowColumnStripes!.Value);
            Assert.True(prices.TableStyleInfo.ShowFirstColumn!.Value);
            Assert.NotEqual(sales.Id!.Value, prices.Id!.Value);

            // Headers fixed for Excel: empty → ColumnN, repeated → numbered, number → text.
            var raw = TableOf(document, "Raw_data").Table;
            Assert.Equal(new[] { "Name", "Column2", "Name2", "2026" }, raw.TableColumns!.Elements<TableColumn>().Select(c => c.Name!.Value).ToArray());
            Assert.Equal("Column2", TextOf(document, CellOf(document, "Raw", "B1")));
            Assert.Equal("Name2", TextOf(document, CellOf(document, "Raw", "C1")));
            Assert.Equal(CellValues.SharedString, CellOf(document, "Raw", "D1").DataType!.Value);
            Assert.Null(raw.TableStyleInfo!.Name);
            Assert.True(raw.TotalsRowShown != null && !raw.TotalsRowShown.Value);
            Assert.Equal(3u, raw.Id!.Value);

            // The sheet lists its tables; the ids lead to the parts.
            var salesSheet = (WorksheetPart)document.WorkbookPart!.GetPartById(document.WorkbookPart.Workbook.Sheets!.Elements<Sheet>().First().Id!);
            Assert.Equal(2u, salesSheet.Worksheet.GetFirstChild<TableParts>()!.Count!.Value);
            Assert.Equal(2, salesSheet.TableDefinitionParts.Count());
        }
    }

    [Fact]
    public void SavingTwiceGivesTheSameBytes()
    {
        var folder = XlsxTestKit.NewFolder();
        var book = Sales();
        book.Created = XlsxTestKit.FixedTime;
        book.Save(Path.Combine(folder, "a.xlsx"));
        book.Save(Path.Combine(folder, "b.xlsx"));
        var again = Sales();
        again.Created = XlsxTestKit.FixedTime;
        again.Save(Path.Combine(folder, "c.xlsx"));
        var a = File.ReadAllBytes(Path.Combine(folder, "a.xlsx"));
        Assert.Equal(a, File.ReadAllBytes(Path.Combine(folder, "b.xlsx")));
        Assert.Equal(a, File.ReadAllBytes(Path.Combine(folder, "c.xlsx")));
    }

    [Theory]
    [InlineData(XlsxTotal.Sum, "SUBTOTAL(109,T[V])", "10")]
    [InlineData(XlsxTotal.Average, "SUBTOTAL(101,T[V])", "2.5")]
    [InlineData(XlsxTotal.Count, "SUBTOTAL(103,T[V])", "5")]
    [InlineData(XlsxTotal.CountNumbers, "SUBTOTAL(102,T[V])", "4")]
    [InlineData(XlsxTotal.Min, "SUBTOTAL(105,T[V])", "1")]
    [InlineData(XlsxTotal.Max, "SUBTOTAL(104,T[V])", "4")]
    [InlineData(XlsxTotal.Var, "SUBTOTAL(110,T[V])", "1.6666666666666667")]
    public void TotalsAreComputedLikeExcel(XlsxTotal total, string formula, string result)
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.Set(1, 1, "V");
        for (var i = 1; i <= 4; i++)
        {
            sheet.Set(i + 1, 1, i);
        }

        sheet.Set(6, 1, "text");
        sheet.AddTable(1, 1, 6, 1, "T").Total(1, total);
        var path = Save(book, "total-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var cell = CellOf(document, "One", "A7");
            Assert.Equal(formula, cell.CellFormula!.Text);
            Assert.Equal(double.Parse(result, CultureInfo.InvariantCulture), double.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture), 12);
        }
    }

    [Fact]
    public void TotalsWithoutEnoughNumbersLeaveTheResultToExcel()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.Set(1, 1, "V");
        sheet.Set(1, 2, "W");
        sheet.Set(2, 1, "only text");
        sheet.Set(2, 2, 5.0);
        var table = sheet.AddTable(1, 1, 2, 2, "T");
        table.Total(1, XlsxTotal.Average);
        table.Total(2, XlsxTotal.StdDev);
        var path = Save(book, "total-empty-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            Assert.Null(CellOf(document, "One", "A3").CellValue);
            Assert.Null(CellOf(document, "One", "B3").CellValue);
            Assert.NotNull(CellOf(document, "One", "B3").CellFormula);
        }
    }

    [Fact]
    public void BadTablesAreRejectedWhenAdded()
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        Assert.Throws<ArgumentException>(() => sheet.AddTable(1, 1, 1, 3, "HeaderOnly"));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(1, 1, 5, 3, "Two words"));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(1, 1, 5, 3, "A1"));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(5, 1, 1, 3, "Reversed"));
        var table = sheet.AddTable(1, 1, 5, 3, "First");
        Assert.Throws<ArgumentException>(() => sheet.AddTable(5, 3, 9, 4, "Second"));
        sheet.AddTable(1, 4, 5, 5, "Beside");
        Assert.Throws<ArgumentException>(() => table.Style = "TableStyleMedium29");
        Assert.Throws<ArgumentException>(() => table.Style = "tablestylemedium2");
        Assert.Throws<ArgumentException>(() => table.Style = "TableStyleLight01");
        table.Style = "TableStyleDark11";
        Assert.Throws<ArgumentOutOfRangeException>(() => table.Total(4, XlsxTotal.Sum));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.TotalLabel(0, "x"));
    }

    [Fact]
    public void BadTablesAreRejectedWhenSaved()
    {
        void Fails(Action<XlsxWorkbook> build)
        {
            var book = new XlsxWorkbook();
            build(book);
            var folder = XlsxTestKit.NewFolder();
            Assert.Throws<InvalidOperationException>(() => book.Save(Path.Combine(folder, "bad.xlsx")));
            Assert.Empty(Directory.GetFiles(folder));
        }

        // Names are unique across sheets, ignoring case, and against defined names.
        Fails(b =>
        {
            b.AddSheet("One").AddTable(1, 1, 3, 2, "Data");
            b.AddSheet("Two").AddTable(1, 1, 3, 2, "DATA");
        });
        Fails(b =>
        {
            var s = b.AddSheet("One");
            s.AddTable(1, 1, 3, 2, "Data");
            b.DefineName("data", s, 5, 5, 6, 6);
        });

        // Merged cells and the sheet's auto filter do not mix with tables.
        Fails(b =>
        {
            var s = b.AddSheet("One");
            s.AddTable(1, 1, 3, 2, "Data");
            s.Merge(2, 1, 2, 2);
        });
        Fails(b =>
        {
            var s = b.AddSheet("One");
            s.AddTable(1, 1, 3, 2, "Data");
            s.AutoFilter(3, 2, 8, 3);
        });

        // A totals row switched on later must not run into the table below.
        Fails(b =>
        {
            var s = b.AddSheet("One");
            var upper = s.AddTable(1, 1, 3, 2, "Upper");
            s.AddTable(4, 1, 6, 2, "Lower");
            upper.Total(2, XlsxTotal.Sum);
        });
    }

    [Fact]
    public void ATableNextToAMergeAndAFilterIsFine()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.AddTable(1, 1, 3, 2, "Data");
        sheet.Merge(5, 1, 5, 2);
        sheet.AutoFilter(1, 4, 3, 5);
        Save(book, "neighbours-check.xlsx");
    }

    [Fact]
    public void FormulaHeadersBecomeText()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.SetFormula(1, 1, "\"Made\"&\" up\"", "Made up");
        sheet.SetRichText(1, 2, new XlsxRun("Rich", new XlsxStyle { Bold = true }));
        sheet.Set(2, 1, 1.0);
        sheet.AddTable(1, 1, 2, 2, "T");
        var path = Save(book, "formula-header-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var a1 = CellOf(document, "One", "A1");
            Assert.Null(a1.CellFormula);
            Assert.Equal("Made up", TextOf(document, a1));
            Assert.Equal("Rich", TextOf(document, CellOf(document, "One", "B1")));
            Assert.Equal(new[] { "Made up", "Rich" }, TableOf(document, "T").Table.TableColumns!.Elements<TableColumn>().Select(c => c.Name!.Value).ToArray());
        }
    }

    [Fact]
    public void AValueReplacesAFormula()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.SetFormula(1, 1, "1+1");
        sheet.Set(1, 1, "text");
        sheet.SetFormula(1, 2, "1+1");
        sheet.Set(1, 2, 5.0);
        sheet.SetFormula(1, 3, "1+1");
        sheet.Set(1, 3, true);
        sheet.SetFormula(1, 4, "1+1");
        sheet.Set(1, 4, new DateTime(2026, 1, 1));
        var path = Save(book, "value-replace-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            Assert.All(document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>(), c => Assert.Null(c.CellFormula));
            Assert.Null(document.WorkbookPart.Workbook.CalculationProperties);
        }
    }

    [Theory]
    [InlineData("Plain", "Plain")]
    [InlineData("Price [USD]", "Price '[USD']")]
    [InlineData("#Items", "'#Items")]
    [InlineData("Owner's", "Owner''s")]
    public void ColumnNamesAreEscapedInReferences(string name, string escaped) => Assert.Equal(escaped, XlsxTable.Escape(name));
}
