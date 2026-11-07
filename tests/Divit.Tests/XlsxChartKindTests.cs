namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

/// <summary>
/// Column, bar, line, area, pie and doughnut charts: the chart element of each kind, its grouping, axes, colours,
/// categories and labels.
/// </summary>
public class XlsxChartKindTests
{
    private static readonly string[] Products = { "Pens", "Paper", "Ink" };

    /// <summary>
    /// Monthly sales of three products, and a chart of every kind beside them.
    /// </summary>
    private static XlsxWorkbook Charts()
    {
        var book = new XlsxWorkbook { Title = "Chart kinds" };
        var sheet = book.AddSheet("Sales");
        sheet.Set(1, 1, "Month");
        for (var p = 0; p < Products.Length; p++)
        {
            sheet.Set(1, p + 2, Products[p]);
        }

        sheet.Set(1, 5, "Start of month");
        sheet.Format(1, 1, 1, 5, (ref XlsxStyle s) => s.Bold = true);
        var months = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat.AbbreviatedMonthNames;
        for (var m = 0; m < 12; m++)
        {
            var row = m + 2;
            sheet.Set(row, 1, months[m]);
            sheet.Set(row, 2, 120 + (m * 9) + ((m % 3) * 14));
            sheet.Set(row, 3, 80 + (m * 4) + ((m % 4) * 10));
            sheet.Set(row, 4, 40 + ((m * 7) % 30));
            sheet.Set(row, 5, new DateTime(2026, m + 1, 1));
            sheet.Format(row, 5, (ref XlsxStyle s) => s.NumberFormat = "mmm yyyy");
        }

        sheet.ColumnWidth(5, 14);
        var width = 520;
        var height = 300;

        var column = sheet.AddChart(1, 7, width, height);
        column.Kind = XlsxChartKind.Column;
        column.Title = "Column, side by side";
        column.YAxis.Title = "Units";
        column.AddSeriesFromColumns(1, 13, 1, 2, 4);

        var stacked = sheet.AddChart(1, 16, width, height);
        stacked.Kind = XlsxChartKind.Column;
        stacked.Grouping = XlsxChartGrouping.Stacked;
        stacked.GapWidth = 60;
        stacked.Title = "Column, stacked";
        stacked.AddSeriesFromColumns(1, 13, 1, 2, 4);

        var bar = sheet.AddChart(17, 7, width, height);
        bar.Kind = XlsxChartKind.Bar;
        bar.Grouping = XlsxChartGrouping.PercentStacked;
        bar.Title = "Bar, stacked to 100 %";
        bar.AddSeriesFromColumns(1, 7, 1, 2, 4);

        var line = sheet.AddChart(17, 16, width, height);
        line.Kind = XlsxChartKind.Line;
        line.Title = "Line over dates";
        line.XAxis.Title = "Month";
        line.AddSeries("Pens", 2, 13, 5, 2).MarkerSize = 6;
        line.AddSeries("Target", 2, 13, 5, 3, new XlsxLine("#C62828", 1.5, "dash"));

        var area = sheet.AddChart(33, 7, width, height);
        area.Kind = XlsxChartKind.Area;
        area.Grouping = XlsxChartGrouping.Stacked;
        area.Title = "Area, stacked";
        area.AddSeriesFromColumns(1, 13, 1, 2, 4);

        var labelled = sheet.AddChart(33, 16, width, height);
        labelled.Kind = XlsxChartKind.Column;
        labelled.Title = "Column with values";
        labelled.DataLabels = XlsxDataLabels.Value;
        labelled.YAxis.Gridlines = false;
        labelled.AddSeries("Ink", 2, 7, 1, 4);

        var pie = sheet.AddChart(49, 7, 420, 320);
        pie.Kind = XlsxChartKind.Pie;
        pie.Title = "Pie: share of pens per quarter month";
        pie.DataLabels = XlsxDataLabels.Percent | XlsxDataLabels.Category;
        pie.FirstSliceAngle = 90;
        pie.AddSeries("Pens", 2, 5, 1, 2).Explosion = 8;

        var doughnut = sheet.AddChart(49, 16, 420, 320);
        doughnut.Kind = XlsxChartKind.Doughnut;
        doughnut.Title = "Doughnut: two products";
        doughnut.HoleSize = 45;
        doughnut.AddSeries("Pens", 2, 5, 1, 2);
        doughnut.AddSeries("Paper", 2, 5, 1, 3);
        return book;
    }

    private static C.PlotArea[] Plots(SpreadsheetDocument document) =>
        document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts
            .Select(p => p.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!).ToArray();

    private static SpreadsheetDocument Open(out string path)
    {
        var book = Charts();
        book.Created = XlsxTestKit.FixedTime;
        path = Path.Combine(XlsxTestKit.NewFolder(), "charts.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        return SpreadsheetDocument.Open(path, false);
    }

    [Fact]
    public void EveryKindIsValidAndKept()
    {
        string path;
        using (var document = Open(out path))
        {
            Assert.Equal(8, Plots(document).Length);
        }

        XlsxTestKit.Keep(path, "charts.xlsx");
    }

    [Fact]
    public void ColumnAndBarCharts()
    {
        using (var document = Open(out _))
        {
            var plots = Plots(document);
            var column = plots[0].GetFirstChild<C.BarChart>()!;
            Assert.Equal(C.BarDirectionValues.Column, column.BarDirection!.Val!.Value);
            Assert.Equal(C.BarGroupingValues.Clustered, column.BarGrouping!.Val!.Value);
            Assert.Null(column.GetFirstChild<C.Overlap>());
            Assert.Equal((ushort)150, column.GetFirstChild<C.GapWidth>()!.Val!.Value);
            var series = column.Elements<C.BarChartSeries>().ToList();
            Assert.Equal(new[] { "Pens", "Paper", "Ink" }, series.Select(s => s.SeriesText!.InnerText).ToArray());

            // Colours from the palette, in turn.
            Assert.Equal(new[] { "4472C4", "ED7D31", "A5A5A5" },
                series.Select(s => s.ChartShapeProperties!.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!.RgbColorModelHex!.Val!.Value).ToArray());

            // Text categories: a string reference with its cache.
            var categories = series[0].GetFirstChild<C.CategoryAxisData>()!.StringReference!;
            Assert.Equal("'Sales'!$A$2:$A$13", categories.Formula!.Text);
            Assert.Equal("Jan", categories.StringCache!.Elements<C.StringPoint>().First().NumericValue!.Text);
            Assert.Equal(12u, categories.StringCache.PointCount!.Val!.Value);

            var stacked = plots[1].GetFirstChild<C.BarChart>()!;
            Assert.Equal(C.BarGroupingValues.Stacked, stacked.BarGrouping!.Val!.Value);
            Assert.Equal((sbyte)100, stacked.GetFirstChild<C.Overlap>()!.Val!.Value);
            Assert.Equal((ushort)60, stacked.GetFirstChild<C.GapWidth>()!.Val!.Value);

            // A bar chart turns the axes; 100 % stacking shows percentages.
            var bar = plots[2];
            Assert.Equal(C.BarDirectionValues.Bar, bar.GetFirstChild<C.BarChart>()!.BarDirection!.Val!.Value);
            Assert.Equal(C.BarGroupingValues.PercentStacked, bar.GetFirstChild<C.BarChart>()!.BarGrouping!.Val!.Value);
            Assert.Equal(C.AxisPositionValues.Left, bar.GetFirstChild<C.CategoryAxis>()!.AxisPosition!.Val!.Value);
            var values = bar.GetFirstChild<C.ValueAxis>()!;
            Assert.Equal(C.AxisPositionValues.Bottom, values.AxisPosition!.Val!.Value);
            Assert.Equal("0%", values.NumberingFormat!.FormatCode!.Value);

            // Category axes have no gridlines by default, value axes do unless turned off.
            Assert.Null(plots[0].GetFirstChild<C.CategoryAxis>()!.MajorGridlines);
            Assert.NotNull(plots[0].GetFirstChild<C.ValueAxis>()!.MajorGridlines);
            Assert.Null(plots[5].GetFirstChild<C.ValueAxis>()!.MajorGridlines);
            Assert.Equal("Units", plots[0].GetFirstChild<C.ValueAxis>()!.Title!.InnerText);
        }
    }

    [Fact]
    public void LineAndAreaCharts()
    {
        using (var document = Open(out _))
        {
            var plots = Plots(document);
            var line = plots[3].GetFirstChild<C.LineChart>()!;
            Assert.Equal(C.GroupingValues.Standard, line.Grouping!.Val!.Value);
            var series = line.Elements<C.LineChartSeries>().ToList();

            // Line charts do not smooth unless asked; the marker and own colour are kept.
            Assert.All(series, s => Assert.False(s.GetFirstChild<C.Smooth>()!.Val!.Value));
            Assert.Equal(C.MarkerStyleValues.Circle, series[0].Marker!.Symbol!.Val!.Value);
            Assert.Equal(C.MarkerStyleValues.None, series[1].Marker!.Symbol!.Val!.Value);
            var dash = series[1].ChartShapeProperties!.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()!;
            Assert.Equal("C62828", dash.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash, dash.GetFirstChild<DocumentFormat.OpenXml.Drawing.PresetDash>()!.Val!.Value);

            // Date categories: numbers in the cells' date format.
            var dates = series[0].GetFirstChild<C.CategoryAxisData>()!.NumberReference!;
            Assert.Equal("mmm yyyy", dates.NumberingCache!.FormatCode!.Text);
            Assert.Equal(new DateTime(2026, 1, 1).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture),
                dates.NumberingCache.Elements<C.NumericPoint>().First().NumericValue!.Text);
            Assert.Equal("Month", plots[3].GetFirstChild<C.CategoryAxis>()!.Title!.InnerText);

            var area = plots[4].GetFirstChild<C.AreaChart>()!;
            Assert.Equal(C.GroupingValues.Stacked, area.Grouping!.Val!.Value);
            Assert.Equal(3, area.Elements<C.AreaChartSeries>().Count());

            var labels = plots[5].GetFirstChild<C.BarChart>()!.GetFirstChild<C.DataLabels>()!;
            Assert.True(labels.GetFirstChild<C.ShowValue>()!.Val!.Value);
            Assert.False(labels.GetFirstChild<C.ShowPercent>()!.Val!.Value);
            Assert.False(labels.GetFirstChild<C.ShowCategoryName>()!.Val!.Value);
        }
    }

    [Fact]
    public void PieAndDoughnutCharts()
    {
        using (var document = Open(out _))
        {
            var parts = document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.ToList();
            var plots = Plots(document);
            var pie = plots[6].GetFirstChild<C.PieChart>()!;
            Assert.True(pie.VaryColors!.Val!.Value);
            Assert.Equal((ushort)90, pie.GetFirstChild<C.FirstSliceAngle>()!.Val!.Value);
            Assert.Null(plots[6].GetFirstChild<C.ValueAxis>());
            var series = pie.GetFirstChild<C.PieChartSeries>()!;
            Assert.Equal(8u, series.Explosion!.Val!.Value);
            var points = series.Elements<C.DataPoint>().ToList();
            Assert.Equal(4, points.Count);
            Assert.Equal("ED7D31", points[1].ChartShapeProperties!.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            var labels = pie.GetFirstChild<C.DataLabels>()!;
            Assert.True(labels.GetFirstChild<C.ShowPercent>()!.Val!.Value);
            Assert.True(labels.GetFirstChild<C.ShowCategoryName>()!.Val!.Value);
            Assert.False(labels.GetFirstChild<C.ShowValue>()!.Val!.Value);
            Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()!.Val!.Value);

            // A pie chart shows its legend even with one series.
            Assert.NotNull(parts[6].ChartSpace.GetFirstChild<C.Chart>()!.Legend);
            Assert.Null(parts[5].ChartSpace.GetFirstChild<C.Chart>()!.Legend);

            var doughnut = plots[7].GetFirstChild<C.DoughnutChart>()!;
            Assert.Equal((byte)45, doughnut.GetFirstChild<C.HoleSize>()!.Val!.Value);
            Assert.Equal(2, doughnut.Elements<C.PieChartSeries>().Count());
        }
    }

    [Fact]
    public void ScatterChartsKeepTheirSmoothLines()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        for (var r = 1; r <= 5; r++)
        {
            sheet.Set(r, 1, r);
            sheet.Set(r, 2, r * r);
        }

        var chart = sheet.AddChart(1, 4, 300, 200);
        chart.AddSeries("Squares", 1, 5, 1, 2);
        chart.DataLabels = XlsxDataLabels.Value | XlsxDataLabels.SeriesName;
        var path = Path.Combine(XlsxTestKit.NewFolder(), "scatter-labels.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var scatter = Plots(document)[0].GetFirstChild<C.ScatterChart>()!;
            Assert.True(scatter.GetFirstChild<C.ScatterChartSeries>()!.GetFirstChild<C.Smooth>()!.Val!.Value);
            Assert.True(scatter.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowSeriesName>()!.Val!.Value);
        }
    }

    [Fact]
    public void SeriesFromColumnsTakeTheirNamesFromTheHeader()
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        sheet.Set(1, 2, "First");
        var chart = sheet.AddChart(1, 5, 300, 200);
        chart.Kind = XlsxChartKind.Column;
        var series = chart.AddSeriesFromColumns(1, 4, 1, 2, 3);
        Assert.Equal(new[] { "First", "C" }, series.Select(s => s.Name).ToArray());
        Assert.All(series, s => Assert.Equal(2, s.FirstRow));
        Assert.Throws<ArgumentException>(() => chart.AddSeriesFromColumns(1, 4, 1, 3, 2));
    }

    [Fact]
    public void SettingsOutsideTheirRangeAreClamped()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.Set(1, 1, "a");
        sheet.Set(1, 2, 1.0);
        var doughnut = sheet.AddChart(1, 4, 300, 200);
        doughnut.Kind = XlsxChartKind.Doughnut;
        doughnut.HoleSize = 99;
        doughnut.FirstSliceAngle = 500;
        doughnut.AddSeries("S", 1, 1, 1, 2).Explosion = 1000;
        var column = sheet.AddChart(12, 4, 300, 200);
        column.Kind = XlsxChartKind.Column;
        column.GapWidth = -5;
        column.AddSeries("S", 1, 1, 1, 2);
        var path = Path.Combine(XlsxTestKit.NewFolder(), "clamp.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var plots = Plots(document);
            var ring = plots[0].GetFirstChild<C.DoughnutChart>()!;
            Assert.Equal((byte)90, ring.GetFirstChild<C.HoleSize>()!.Val!.Value);
            Assert.Equal((ushort)360, ring.GetFirstChild<C.FirstSliceAngle>()!.Val!.Value);
            Assert.Equal(400u, ring.GetFirstChild<C.PieChartSeries>()!.Explosion!.Val!.Value);
            Assert.Equal((ushort)0, plots[1].GetFirstChild<C.BarChart>()!.GetFirstChild<C.GapWidth>()!.Val!.Value);
        }
    }
}
