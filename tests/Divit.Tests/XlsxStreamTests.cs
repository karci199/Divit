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
/// Streamed sheets: rows written while saving, never kept. They must give the same cells as stored sheets, keep
/// memory flat, and refuse what needs the cells in memory.
/// </summary>
public class XlsxStreamTests
{
    private static readonly XlsxStyle Bold = new XlsxStyle { Bold = true, Fill = "#DCE6F1" };
    private static readonly XlsxStyle Money = new XlsxStyle { NumberFormat = "#,##0.00" };

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

    private static Worksheet SheetOf(SpreadsheetDocument document, string name)
    {
        var part = document.WorkbookPart!;
        var id = part.Workbook.Sheets!.Elements<Sheet>().Single(s => s.Name == name).Id!.Value!;
        return ((WorksheetPart)part.GetPartById(id)).Worksheet;
    }

    /// <summary>
    /// A measurement log: header, rows of every kind of value, and the sheet settings that work on streamed sheets.
    /// </summary>
    private static XlsxWorkbook Log(int rows, bool inline = false)
    {
        var book = new XlsxWorkbook { Title = "Streamed sheet" };
        var sheet = book.AddStreamedSheet("Log", w =>
        {
            w.StartRow(height: 22);
            w.Cell("Time", Bold).Cell("Sensor", Bold).Cell("Value", Bold).Cell("Cost", Bold).Cell("OK", Bold).Cell("Note", Bold);
            w.EndRow();
            for (var i = 0; i < rows; i++)
            {
                w.Cell(new DateTime(2026, 1, 1).AddMinutes(i * 15))
                    .Cell("S-" + (i % 4).ToString(CultureInfo.InvariantCulture))
                    .Cell(20 + Math.Sin(i / 10.0))
                    .Cell(i * 0.25, Money)
                    .Cell(i % 7 != 0);
                if (i % 50 == 0)
                {
                    w.RichText("check ", new XlsxRun("now", new XlsxStyle { Bold = true, FontColor = "#C00000" }));
                }

                w.EndRow();
            }

            w.SkipRows(1);
            w.Cell("Total", Bold).Skip(2).Formula("SUM(D2:D" + (rows + 1).ToString(CultureInfo.InvariantCulture) + ")", null, Money);
            w.EndRow();
        }, inline);

        sheet.FrozenRows = 1;
        sheet.ColumnWidth(1, 18);
        sheet.ColumnWidth(6, 14);
        sheet.AutoFilter(1, 1, rows + 1, 6);
        sheet.Highlight(2, 3, rows + 1, 3, XlsxRule.Compare(XlsxComparison.GreaterThan, 20.9), new XlsxStyle { Fill = "#FFC7CE" });
        sheet.DataBar(2, 4, rows + 1, 4, "#638EC6");
        sheet.Validate(2, 2, rows + 1, 2, XlsxValidation.List("S-0", "S-1", "S-2", "S-3"));
        sheet.Comment(1, 3, "Degrees Celsius.");
        sheet.Merge(rows + 3, 1, rows + 3, 2);
        book.AddSheet("Summary").Set(1, 1, "A stored sheet beside the streamed one.");
        return book;
    }

    [Fact]
    public void AStreamedSheetIsValidWithEverySetting()
    {
        var path = Save(Log(10000), "streamed.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sheet = SheetOf(document, "Log");
            Assert.Null(sheet.SheetDimension);
            var rows = sheet.Descendants<Row>().ToList();
            Assert.Equal(10002, rows.Count);
            Assert.Equal(22, rows[0].Height!.Value);
            Assert.Equal(10003u, rows.Last().RowIndex!.Value);

            var cells = rows[1].Elements<Cell>().ToList();
            Assert.Equal(new[] { "A2", "B2", "C2", "D2", "E2", "F2" }, cells.Select(c => c.CellReference!.Value).ToArray());
            Assert.Equal(new DateTime(2026, 1, 1).ToOADate(), double.Parse(cells[0].CellValue!.Text, CultureInfo.InvariantCulture));
            Assert.Equal(CellValues.SharedString, cells[1].DataType!.Value);
            Assert.Equal(CellValues.Boolean, cells[4].DataType!.Value);
            Assert.Equal("0", cells[4].CellValue!.Text);

            var total = rows.Last().Elements<Cell>().ToList();
            Assert.Equal(new[] { "A10003", "D10003" }, total.Select(c => c.CellReference!.Value).ToArray());
            Assert.Equal("SUM(D2:D10001)", total[1].CellFormula!.Text);
            Assert.True(document.WorkbookPart!.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);

            // Shared strings hold each text once: six headers, four sensors, the total label, one rich text and the
            // text of the stored sheet.
            Assert.Equal(13u, document.WorkbookPart.SharedStringTablePart!.SharedStringTable.UniqueCount!.Value);
            Assert.NotNull(sheet.GetFirstChild<AutoFilter>());
            Assert.NotNull(sheet.GetFirstChild<DataValidations>());
            Assert.Single(sheet.Elements<ConditionalFormatting>(), c => c.Descendants<DataBar>().Any());
        }
    }

    [Fact]
    public void InlineStringsKeepTextsInTheirCells()
    {
        var path = Save(Log(100, inline: true), "streamed-inline.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sheet = SheetOf(document, "Log");
            var text = sheet.Descendants<Cell>().Single(c => c.CellReference == "B2");
            Assert.Equal(CellValues.InlineString, text.DataType!.Value);
            Assert.Equal("S-0", text.InlineString!.Text!.Text);
            var rich = sheet.Descendants<Cell>().Single(c => c.CellReference == "F2");
            Assert.Equal(2, rich.InlineString!.Elements<Run>().Count());

            // The stored sheet still uses the shared strings.
            Assert.Contains(document.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>(),
                s => s.InnerText.StartsWith("A stored sheet", StringComparison.Ordinal));
        }
    }

    /// <summary>
    /// The same cells through Set and through the row writer give the same sheetData, byte for byte.
    /// </summary>
    [Fact]
    public void StreamedCellsEqualStoredCells()
    {
        var values = Enumerable.Range(0, 300).Select(i => (Text: "t" + (i % 13), Number: i * 1.5 - 40, Date: new DateTime(2026, 3, 1).AddHours(i), Flag: i % 2 == 0)).ToList();

        var stored = new XlsxWorkbook();
        var sheet = stored.AddSheet("Data");
        for (var i = 0; i < values.Count; i++)
        {
            var row = i + 1;
            sheet.Set(row, 1, values[i].Text);
            sheet.Set(row, 2, values[i].Number);
            sheet.Format(row, 2, (ref XlsxStyle s) => s.NumberFormat = "#,##0.00");
            sheet.Set(row, 3, values[i].Date);
            sheet.Set(row, 4, values[i].Flag);
            sheet.SetFormula(row, 6, "B" + row + "*2", values[i].Number * 2);
        }

        var streamed = new XlsxWorkbook();
        streamed.AddStreamedSheet("Data", w =>
        {
            for (var i = 0; i < values.Count; i++)
            {
                w.Cell(values[i].Text).Cell(values[i].Number, Money).Cell(values[i].Date).Cell(values[i].Flag).Skip()
                    .Formula("B" + w.Row + "*2", values[i].Number * 2);
                w.EndRow();
            }
        });

        string SheetData(XlsxWorkbook book, string name)
        {
            using (var document = SpreadsheetDocument.Open(Save(book, name), false))
            {
                return SheetOf(document, "Data").GetFirstChild<SheetData>()!.OuterXml;
            }
        }

        Assert.Equal(SheetData(stored, "stored-check.xlsx"), SheetData(streamed, "streamed-check.xlsx"));
    }

    [Fact]
    public void RowsCanBeSkippedAndShaped()
    {
        var book = new XlsxWorkbook();
        book.AddStreamedSheet("One", w =>
        {
            w.SkipRows(2);
            w.StartRow(height: 30, hidden: true, outlineLevel: 2);
            w.Skip(3).Cell(1.0).Cell((string?)null, Bold).Cell((string?)null);
            w.EndRow();
            w.WriteRow("a", 2, 3L, 4.5m, 5f, (short)6, (byte)7, null, true, new DateTime(2026, 1, 1));
        });
        var path = Save(book, "skip-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var rows = SheetOf(document, "One").Descendants<Row>().ToList();
            Assert.Equal(3u, rows[0].RowIndex!.Value);
            Assert.True(rows[0].Hidden!.Value);
            Assert.Equal((byte)2, rows[0].OutlineLevel!.Value);
            Assert.Equal(new[] { "D3", "E3" }, rows[0].Elements<Cell>().Select(c => c.CellReference!.Value).ToArray());
            var last = rows[1].Elements<Cell>().ToList();
            Assert.Equal(9, last.Count);
            Assert.Equal("I4", last[7].CellReference!.Value);
            Assert.Equal("4.5", last[3].CellValue!.Text);
        }
    }

    [Fact]
    public void StreamedSheetsRefuseCellAndRowCalls()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddStreamedSheet("One", w => { });
        Assert.True(sheet.IsStreamed);
        Assert.Throws<InvalidOperationException>(() => sheet.Set(1, 1, "x"));
        Assert.Throws<InvalidOperationException>(() => sheet.Set(1, 1, 1.0));
        Assert.Throws<InvalidOperationException>(() => sheet.Format(1, 1, (ref XlsxStyle s) => s.Bold = true));
        Assert.Throws<InvalidOperationException>(() => sheet.SetFormula(1, 1, "1+1"));
        Assert.Throws<InvalidOperationException>(() => sheet.RowHeight(1, 20));
        Assert.Throws<InvalidOperationException>(() => sheet.HideRow(1));
        Assert.Throws<InvalidOperationException>(() => sheet.GroupRows(1, 3));
        Assert.Throws<InvalidOperationException>(() => sheet.AddTable(1, 1, 3, 2, "T"));

        // Column settings, merges, links between sheets and comments stay available.
        sheet.ColumnWidth(1, 20);
        sheet.HideColumn(2);
        sheet.Merge(1, 1, 1, 3);
        sheet.Link(2, 1, "One");
        sheet.Comment(1, 1, "note");
        Assert.False(book.AddSheet("Two").IsStreamed);
        Save(book, "refuse-check.xlsx");
    }

    [Fact]
    public void RowWriterChecksItsUse()
    {
        void Fails<T>(Action<XlsxRowWriter> write)
            where T : Exception
        {
            var book = new XlsxWorkbook();
            book.AddStreamedSheet("One", write);
            var folder = XlsxTestKit.NewFolder();
            Assert.Throws<T>(() => book.Save(Path.Combine(folder, "bad.xlsx")));
            Assert.Empty(Directory.GetFiles(folder));
        }

        Fails<InvalidOperationException>(w => { w.StartRow(); w.StartRow(); });
        Fails<ArgumentOutOfRangeException>(w => w.Skip(16384).Cell(1.0).Cell(2.0));
        Fails<ArgumentOutOfRangeException>(w => { w.SkipRows(1048576); w.Cell(1.0); });
        Fails<ArgumentException>(w => w.WriteRow(new object()));
        Fails<ArgumentException>(w => w.Formula(" = "));
        Fails<ArgumentOutOfRangeException>(w => w.Cell(new DateTime(1800, 1, 1)));

        // The last row and the last column of Excel are fine.
        var ok = new XlsxWorkbook();
        ok.AddStreamedSheet("One", w =>
        {
            w.Skip(16383).Cell("XFD");
            w.SkipRows(1048574);
            w.Cell("last row");
        });
        var path = Save(ok, "limits-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var cells = SheetOf(document, "One").Descendants<Cell>().Select(c => c.CellReference!.Value).ToArray();
            Assert.Equal(new[] { "XFD1", "A1048576" }, cells);
        }
    }

    [Fact]
    public void EverySaveRunsTheCallbackAndGivesTheSameBytes()
    {
        var calls = 0;
        var book = new XlsxWorkbook { Created = XlsxTestKit.FixedTime };
        book.AddStreamedSheet("One", w =>
        {
            calls++;
            for (var i = 0; i < 100; i++)
            {
                w.Cell("row").Cell(i).EndRow();
            }
        });
        var folder = XlsxTestKit.NewFolder();
        book.Save(Path.Combine(folder, "a.xlsx"));
        book.Save(Path.Combine(folder, "b.xlsx"));
        Assert.Equal(2, calls);
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.xlsx")), File.ReadAllBytes(Path.Combine(folder, "b.xlsx")));
    }

#if NET
    /// <summary>
    /// A million number cells with a style allocate almost nothing per cell: nothing is kept, nothing is boxed.
    /// </summary>
    [Fact]
    public void AMillionCellsKeepMemoryFlat()
    {
        var book = new XlsxWorkbook();
        book.AddStreamedSheet("Numbers", w =>
        {
            for (var r = 0; r < 100000; r++)
            {
                for (var c = 0; c < 10; c++)
                {
                    w.Cell(r * 0.5 + c, Money);
                }

                w.EndRow();
            }
        });
        var path = Path.Combine(XlsxTestKit.NewFolder(), "million.xlsx");
        book.Save(path);
        var before = GC.GetAllocatedBytesForCurrentThread();
        book.Save(path);
        var perCell = (GC.GetAllocatedBytesForCurrentThread() - before) / 1000000.0;
        Assert.True(perCell < 8, "allocated " + perCell.ToString("0.0", CultureInfo.InvariantCulture) + " bytes per cell");
    }
#endif
}
