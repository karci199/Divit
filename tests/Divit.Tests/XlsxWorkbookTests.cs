namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Xunit;

/// <summary>
/// The xlsx writer: every feature in one valid file, reproducible output, culture independence, messages and safe saving.
/// </summary>
public class XlsxWorkbookTests
{
    /// <summary>
    /// A workbook that uses every feature of the writer.
    /// </summary>
    private static XlsxWorkbook FullWorkbook()
    {
        var book = new XlsxWorkbook { Title = "Divit sample", Subject = "All features", Keywords = "SN-001; sample", Created = XlsxTestKit.FixedTime };
        var data = book.AddSheet("Data");
        data.TabColor = "#1F4E9A";
        data.FrozenRows = 1;
        data.Zoom = 90;
        data.Landscape = true;
        data.Header = "&LDivit&RPage &P / &N";
        data.PrintTitleRows = (1, 1);
        data.PrintArea = (1, 1, 40, 4);
        data.Set(1, 1, "Angle");
        data.Set(1, 2, "Gain (dB)");
        data.Set(1, 3, "Limit (dB)");
        data.Set(1, 4, "Serial");
        data.Format(1, 1, 1, 4, (ref XlsxStyle s) => { s.Bold = true; s.Fill = "#DCE6F1"; s.Horizontal = "center"; });
        for (var i = 0; i < 37; i++)
        {
            var angle = i * 10.0;
            data.Set(i + 2, 1, angle);
            data.Set(i + 2, 2, 10 * Math.Cos(angle * Math.PI / 180) - 3);
            data.Set(i + 2, 3, 2.5);
            data.Set(i + 2, 4, "0042");
        }

        data.Format(2, 2, 38, 3, (ref XlsxStyle s) => s.NumberFormat = "0.00");
        data.Frame(1, 1, 38, 4, new XlsxBorder("medium", "#404040"), new XlsxBorder("thin", "#BFBFBF"));
        data.ColumnWidth(2, 14);
        data.RowHeight(1, 22);
        data.AutoFilter(1, 1, 38, 4);
        data.ColorScale(2, 2, 38, 2, "#F8696B", "#FFEB84", "#63BE7B");
        data.IgnoreNumbersAsText = true;
        data.Set(40, 1, double.NaN);

        var chart = data.AddChart(2, 6, 640, 360);
        chart.Title = "Gain";
        chart.Subtitle = "Phi = 0°";
        chart.XAxis.Title = "Angle (°)";
        chart.YAxis.Title = "dB";
        chart.XAxis.Min = 0;
        chart.XAxis.Max = 360;
        chart.XAxis.MajorUnit = 30;
        chart.AddSeries("Gain", 2, 38, 1, 2);
        chart.AddSeries("Limit", 2, 38, 1, 3, chart.Look.AccentLine).Smooth = false;
        chart.AddPoints("Marked", 2, 5, 1, 2);

        var polar = data.AddChart(22, 6, 420, 420);
        polar.Kind = XlsxChartKind.Polar;
        polar.Title = "Pattern";
        polar.XAxis.MajorUnit = 30;
        polar.YAxis.Min = -15;
        polar.AddSeries("Gain", 2, 37, 1, 2);

        var summary = book.AddSheet("Summary");
        summary.Set(1, 1, "Overview");
        summary.Merge(1, 1, 1, 3);
        summary.Link(2, 1, "Data");
        summary.Set(2, 1, "Go to data");
        summary.LinkFile(3, 1, "Other file", "other.xlsx", "Sheet 1");
        summary.HideColumn(5);
        summary.ShowGridLines = false;
        summary.AddImage(5, 1, OnePixelPng(), 32, 32);
        summary.Protect("secret");
        return book;
    }

    /// <summary>
    /// The smallest valid PNG: one transparent pixel.
    /// </summary>
    private static byte[] OnePixelPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    [Fact]
    public void FullWorkbookIsValid()
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), "full.xlsx");
        FullWorkbook().Save(path);

        Assert.Empty(XlsxTestKit.Validate(path));
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sheets = document.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Select(s => s.Name!.Value).ToArray();
            Assert.Equal(new[] { "Data", "Summary" }, sheets);
            Assert.Equal("Divit sample", document.PackageProperties.Title);
        }

        XlsxTestKit.Keep(path, "full.xlsx");
    }

    [Fact]
    public void SameInputGivesSameBytes()
    {
        var folder = XlsxTestKit.NewFolder();
        var a = Path.Combine(folder, "a.xlsx");
        var b = Path.Combine(folder, "b.xlsx");
        FullWorkbook().Save(a);
        FullWorkbook().Save(b);

        Assert.Equal(File.ReadAllBytes(a), File.ReadAllBytes(b));
    }

    [Fact]
    public void NumbersIgnoreTheThreadCulture()
    {
        var old = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("tr-TR");
            var book = new XlsxWorkbook();
            var sheet = book.AddSheet("Sayı");
            sheet.Set(1, 1, 1.5);
            sheet.Set(2, 1, "İ");
            var path = Path.Combine(XlsxTestKit.NewFolder(), "culture.xlsx");
            book.Save(path);

            Assert.Empty(XlsxTestKit.Validate(path));
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().First();
                Assert.Equal("1.5", cell.CellValue!.Text);
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = old;
        }
    }

    [Fact]
    public void SavingReplacesTheTargetAndLeavesNoTemporaryFile()
    {
        var folder = XlsxTestKit.NewFolder();
        var path = Path.Combine(folder, "book.xlsx");
        File.WriteAllText(path, "old content");

        var book = new XlsxWorkbook();
        book.AddSheet("One").Set(1, 1, "new");
        book.Save(path);

        Assert.Empty(XlsxTestKit.Validate(path));
        Assert.Equal(new[] { "book.xlsx" }, Directory.GetFiles(folder).Select(Path.GetFileName).ToArray());
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("A name that is longer than thirty one", false)]
    [InlineData("Bad/name", false)]
    [InlineData("'quoted'", false)]
    [InlineData("Good name", true)]
    [InlineData("Ölçüm 1", true)]
    public void SheetNamesAreChecked(string name, bool usable)
    {
        Assert.Equal(usable, new XlsxWorkbook().SheetNameProblem(name) == null);
    }

    [Fact]
    public void DuplicateSheetNamesIgnoreCase()
    {
        var book = new XlsxWorkbook();
        book.AddSheet("Data");
        Assert.NotNull(book.SheetNameProblem("DATA"));
    }

    [Fact]
    public void EmptyWorkbookCannotBeSaved()
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), "empty.xlsx");
        Assert.Throws<InvalidOperationException>(() => new XlsxWorkbook().Save(path));
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void MessagesFollowTheUiLanguage()
    {
        var old = Thread.CurrentThread.CurrentUICulture;
        try
        {
            var sheet = new XlsxWorkbook().AddSheet("One");
            Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture;
            var english = Assert.Throws<ArgumentException>(() => sheet.Merge(3, 3, 1, 1)).Message;
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("tr-TR");
            var turkish = Assert.Throws<ArgumentException>(() => sheet.Merge(3, 3, 1, 1)).Message;

            Assert.StartsWith("Merge range is reversed", english);
            Assert.StartsWith("Birleştirme aralığı ters", turkish);
        }
        finally
        {
            Thread.CurrentThread.CurrentUICulture = old;
        }
    }

    [Fact]
    public void CellsOutsideExcelLimitsAreRejected()
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.Set(0, 1, "x"));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.Set(1, 16385, "x"));
        Assert.Throws<ArgumentException>(() => sheet.AddImage(1, 1, new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }, 10, 10));
    }
}
