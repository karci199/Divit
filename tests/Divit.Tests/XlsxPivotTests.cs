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
/// Pivot tables: the cache, its records, the table definition, and the summary written into the cells.
/// </summary>
public class XlsxPivotTests
{
    private static readonly string[] Regions = { "North", "South", "East" };
    private static readonly string[] Products = { "Pens", "Paper", "Ink" };
    private static readonly string[] Months = { "Jan", "Feb", "Mar", "Apr" };

    /// <summary>
    /// Sixty sales rows, and a summary sheet with three pivot tables over them.
    /// </summary>
    private static XlsxWorkbook Sales()
    {
        var book = new XlsxWorkbook { Title = "Pivot tables", Created = XlsxTestKit.FixedTime };
        var data = book.AddSheet("Sales");
        var headers = new[] { "Month", "Region", "Product", "Quantity", "Amount" };
        for (var c = 0; c < headers.Length; c++)
        {
            data.Set(1, c + 1, headers[c]);
        }

        data.Format(1, 1, 1, 5, (ref XlsxStyle s) => s.Bold = true);
        for (var i = 0; i < 60; i++)
        {
            var row = i + 2;
            data.Set(row, 1, Months[i % 4]);
            data.Set(row, 2, Regions[i % 3]);
            data.Set(row, 3, Products[(i / 3) % 3]);
            data.Set(row, 4, 1 + (i % 7));
            data.Set(row, 5, 10 + ((i * 37) % 90) + 0.5);
        }

        var summary = book.AddSheet("Summary");
        summary.Set(1, 1, "Drag the fields in Excel's PivotTable Fields pane to rearrange the summaries.");
        var byRegion = summary.AddPivotTable(3, 1, data, 1, 1, 61, 5, "SalesByRegion");
        byRegion.RowField = "Region";
        byRegion.ColumnField = "Product";
        byRegion.FilterField = "Month";
        byRegion.ValueField = "Amount";
        byRegion.NumberFormat = "#,##0.00";
        byRegion.Style = "PivotStyleMedium9";

        var count = summary.AddPivotTable(14, 1, data, 1, 1, 61, 5, "OrdersPerProduct");
        count.RowField = "Product";
        count.ValueField = "Region";
        count.Function = XlsxPivotFunction.Count;

        var average = summary.AddPivotTable(14, 5, data, 1, 1, 61, 5, "QuantityByMonth");
        average.RowField = "Month";
        average.ValueField = "quantity";
        average.Function = XlsxPivotFunction.Average;
        summary.ColumnWidth(1, 16);
        return book;
    }

    private static string Save(XlsxWorkbook book, string name)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        return path;
    }

    private static Worksheet SheetOf(SpreadsheetDocument document, string name)
    {
        var part = document.WorkbookPart!;
        var id = part.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == name).Id!.Value!;
        return ((WorksheetPart)part.GetPartById(id)).Worksheet;
    }

    private static object? CellOf(SpreadsheetDocument document, string sheet, string reference)
    {
        var cell = SheetOf(document, sheet).Descendants<Cell>().SingleOrDefault(c => c.CellReference == reference);
        if (cell?.CellValue == null)
        {
            return null;
        }

        return cell.DataType?.Value == CellValues.SharedString
            ? document.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.Text)).InnerText
            : (object)double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);
    }

    [Fact]
    public void PivotTablesAreValidAndLinked()
    {
        var path = Save(Sales(), "pivot.xlsx");
        XlsxTestKit.Keep(path, "pivot.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var workbook = document.WorkbookPart!;
            var caches = workbook.Workbook.GetFirstChild<PivotCaches>()!.Elements<PivotCache>().ToList();
            Assert.Equal(new uint[] { 1, 2, 3 }, caches.Select(c => c.CacheId!.Value).ToArray());
            Assert.All(caches, c => Assert.IsType<PivotTableCacheDefinitionPart>(workbook.GetPartById(c.Id!.Value!)));

            var tables = ((WorksheetPart)workbook.GetPartById(workbook.Workbook.Sheets!.Elements<Sheet>().Last().Id!.Value!)).PivotTableParts.ToList();
            Assert.Equal(3, tables.Count);
            Assert.All(tables, t => Assert.NotNull(t.PivotTableCacheDefinitionPart!.PivotTableCacheRecordsPart));

            var cache = tables[0].PivotTableCacheDefinitionPart!.PivotCacheDefinition;
            Assert.True(cache.RefreshOnLoad!.Value);
            Assert.Equal(60u, cache.RecordCount!.Value);
            Assert.Equal("A1:E61", cache.CacheSource!.WorksheetSource!.Reference!.Value);
            Assert.Equal("Sales", cache.CacheSource.WorksheetSource.Sheet!.Value);
            var fields = cache.CacheFields!.Elements<CacheField>().ToList();
            Assert.Equal(new[] { "Month", "Region", "Product", "Quantity", "Amount" }, fields.Select(f => f.Name!.Value).ToArray());

            // Axis fields list their items in order of first appearance; the value field keeps only its range.
            Assert.Equal(new[] { "North", "South", "East" }, fields[1].SharedItems!.Elements<StringItem>().Select(s => s.Val!.Value).ToArray());
            Assert.Empty(fields[4].SharedItems!.ChildElements);
            Assert.True(fields[4].SharedItems!.ContainsNumber!.Value);
            Assert.Equal(10.5, fields[4].SharedItems!.MinValue!.Value);
            Assert.Null(fields[4].SharedItems!.ContainsInteger);
            Assert.True(fields[3].SharedItems!.ContainsInteger!.Value);

            var records = tables[0].PivotTableCacheDefinitionPart!.PivotTableCacheRecordsPart!.PivotCacheRecords;
            Assert.Equal(60u, records.Count!.Value);
            var first = records.Elements<PivotCacheRecord>().First().ChildElements.ToList();
            Assert.Equal(5, first.Count);
            Assert.IsType<FieldItem>(first[1]);
            Assert.Equal("10.5", ((NumberItem)first[4]).Val!.Value.ToString(CultureInfo.InvariantCulture));
        }
    }

    [Fact]
    public void TheTableDefinitionMatchesTheLayout()
    {
        var path = Save(Sales(), "pivot-layout.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var workbook = document.WorkbookPart!;
            var tables = ((WorksheetPart)workbook.GetPartById(workbook.Workbook.Sheets!.Elements<Sheet>().Last().Id!.Value!)).PivotTableParts
                .Select(p => p.PivotTableDefinition).ToList();
            var byRegion = tables.Single(t => t.Name == "SalesByRegion");

            // A filter puts the body two rows lower: A5, header rows, three regions, the total; columns for three
            // products and the total.
            Assert.Equal("A5:E10", byRegion.Location!.Reference!.Value);
            Assert.Equal(2u, byRegion.Location.FirstDataRow!.Value);
            Assert.Equal(1u, byRegion.Location.RowPageCount!.Value);
            var pivotFields = byRegion.PivotFields!.Elements<PivotField>().ToList();
            Assert.Equal(PivotTableAxisValues.AxisPage, pivotFields[0].Axis!.Value);
            Assert.Equal(PivotTableAxisValues.AxisRow, pivotFields[1].Axis!.Value);
            Assert.Equal(PivotTableAxisValues.AxisColumn, pivotFields[2].Axis!.Value);
            Assert.True(pivotFields[4].DataField!.Value);

            // Items in sorted order (East, North, South), then the default total item.
            Assert.Equal(new uint?[] { 2, 0, 1, null }, pivotFields[1].Items!.Elements<Item>().Select(i => i.Index?.Value).ToArray());
            Assert.Equal(ItemValues.Default, pivotFields[1].Items!.Elements<Item>().Last().ItemType!.Value);
            Assert.Equal(4u, byRegion.RowItems!.Count!.Value);
            Assert.Equal(ItemValues.Grand, byRegion.RowItems.Elements<RowItem>().Last().ItemType!.Value);
            Assert.Equal(1u, byRegion.ColumnFields!.Count!.Value);
            Assert.Equal(0, byRegion.PageFields!.Elements<PageField>().Single().Field!.Value);
            var value = byRegion.DataFields!.Elements<DataField>().Single();
            Assert.Equal("Sum of Amount", value.Name!.Value);
            Assert.Null(value.Subtotal);
            Assert.True(value.NumberFormatId!.Value >= 164);
            Assert.Equal("PivotStyleMedium9", byRegion.PivotTableStyle!.Name!.Value);

            var count = tables.Single(t => t.Name == "OrdersPerProduct");
            Assert.Equal(DataConsolidateFunctionValues.Count, count.DataFields!.Elements<DataField>().Single().Subtotal!.Value);
            Assert.Equal("A14:B18", count.Location!.Reference!.Value);
            Assert.Equal(1u, count.ColumnItems!.Count!.Value);
            Assert.Null(count.ColumnFields);
            Assert.Equal("Average of Quantity", tables.Single(t => t.Name == "QuantityByMonth").DataFields!.Elements<DataField>().Single().Name!.Value);
        }
    }

    [Fact]
    public void TheSummaryIsInTheCellsBeforeExcelRefreshes()
    {
        var path = Save(Sales(), "pivot-cells.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            Assert.Equal("Month", CellOf(document, "Summary", "A3"));
            Assert.Equal("(All)", CellOf(document, "Summary", "B3"));
            Assert.Equal("Sum of Amount", CellOf(document, "Summary", "A5"));
            Assert.Equal("Row Labels", CellOf(document, "Summary", "A6"));
            Assert.Equal(new object?[] { "Ink", "Paper", "Pens", "Grand Total" },
                new[] { "B6", "C6", "D6", "E6" }.Select(r => CellOf(document, "Summary", r)).ToArray());
            Assert.Equal(new object?[] { "East", "North", "South", "Grand Total" },
                new[] { "A7", "A8", "A9", "A10" }.Select(r => CellOf(document, "Summary", r)).ToArray());

            // The grand total is the sum of every amount; each region adds up to it.
            var expected = Enumerable.Range(0, 60).Sum(i => 10 + ((i * 37) % 90) + 0.5);
            Assert.Equal(expected, (double)CellOf(document, "Summary", "E10")!, 6);
            var regions = new[] { "E7", "E8", "E9" }.Sum(r => (double)CellOf(document, "Summary", r)!);
            Assert.Equal(expected, regions, 6);

            // Sorted products: Ink 18 rows, Paper 21, Pens 21.
            Assert.Equal(new object?[] { "Ink", 18.0, "Pens", 21.0 }, new[] { "A15", "B15", "A17", "B17" }.Select(r => CellOf(document, "Summary", r)).ToArray());
            Assert.Equal(60.0, CellOf(document, "Summary", "B18"));
            var januaryQuantity = Enumerable.Range(0, 60).Where(i => i % 4 == 0).Average(i => 1 + (i % 7));
            Assert.Equal(januaryQuantity, (double)CellOf(document, "Summary", "F17")!, 9);
        }
    }

    [Fact]
    public void NumbersOnAnAxisAreListedAndStayNumbers()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("Data");
        sheet.Set(1, 1, "Year");
        sheet.Set(1, 2, "Value");
        for (var i = 0; i < 6; i++)
        {
            sheet.Set(i + 2, 1, 2026 - (i % 3));
            sheet.Set(i + 2, 2, i);
        }

        sheet.Set(8, 2, 9.0);
        var pivot = sheet.AddPivotTable(1, 4, sheet, 1, 1, 8, 2, "ByYear");
        pivot.RowField = "Year";
        pivot.ValueField = "Value";
        pivot.Function = XlsxPivotFunction.Max;
        var path = Save(book, "pivot-numbers.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var cache = document.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotCacheDefinition;
            var year = cache.CacheFields!.Elements<CacheField>().First().SharedItems!;
            Assert.Equal(new[] { 2026.0, 2025.0, 2024.0 }, year.Elements<NumberItem>().Select(n => n.Val!.Value).ToArray());
            Assert.True(year.ContainsBlank!.Value);
            Assert.Single(year.Elements<MissingItem>());
            Assert.Equal(2024.0, CellOf(document, "Data", "D2"));
            Assert.Equal("(blank)", CellOf(document, "Data", "D5"));
            Assert.Equal(9.0, CellOf(document, "Data", "E5"));
            Assert.Equal(9.0, CellOf(document, "Data", "E6"));
        }
    }

    [Fact]
    public void WrongPivotTablesStopTheSave()
    {
        void Fails(Action<XlsxPivotTable> set, bool streamed = false, bool twice = false)
        {
            var book = new XlsxWorkbook();
            XlsxSheet data;
            if (streamed)
            {
                data = book.AddStreamedSheet("Data", w => w.WriteRow("Name", "Value"));
            }
            else
            {
                data = book.AddSheet("Data");
                data.Set(1, 1, "Name");
                data.Set(1, 2, "Value");
                data.Set(2, 1, "a");
                data.Set(2, 2, 1.0);
            }

            var sheet = book.AddSheet("Pivot");
            set(sheet.AddPivotTable(1, 1, data, 1, 1, 2, 2, "P"));
            if (twice)
            {
                set(sheet.AddPivotTable(10, 1, data, 1, 1, 2, 2, "p"));
            }

            var folder = XlsxTestKit.NewFolder();
            Assert.Throws<InvalidOperationException>(() => book.Save(Path.Combine(folder, "bad.xlsx")));
            Assert.Empty(Directory.GetFiles(folder));
        }

        Fails(p => p.ValueField = "Value");
        Fails(p => p.RowField = "Name");
        Fails(p => { p.RowField = "Nmae"; p.ValueField = "Value"; });
        Fails(p => { p.RowField = "Name"; p.ValueField = "Name"; p.Function = XlsxPivotFunction.Count; });
        Fails(p => { p.RowField = "Value"; p.ValueField = "Name"; });
        Fails(p => { p.RowField = "Name"; p.ValueField = "Value"; }, twice: true);
        Fails(p => { p.RowField = "Name"; p.ValueField = "Value"; }, streamed: true);

        var sheetOnly = new XlsxWorkbook().AddSheet("One");
        Assert.Throws<ArgumentException>(() => sheetOnly.AddPivotTable(1, 1, sheetOnly, 1, 1, 1, 2, "P"));
        Assert.Throws<ArgumentException>(() => sheetOnly.AddPivotTable(1, 1, sheetOnly, 1, 1, 3, 2, " "));
    }

    [Fact]
    public void PivotFilesAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        Sales().Save(Path.Combine(folder, "a.xlsx"));
        Sales().Save(Path.Combine(folder, "b.xlsx"));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.xlsx")), File.ReadAllBytes(Path.Combine(folder, "b.xlsx")));
    }
}
