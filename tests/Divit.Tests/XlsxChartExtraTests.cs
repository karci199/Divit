namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

/// <summary>
/// Combined charts, the right-hand axis, logarithmic and reversed axes, trendlines and error bars.
/// </summary>
public class XlsxChartExtraTests
{
    /// <summary>
    /// Monthly sales with a share column, measurements with uncertainties, and the charts built on them.
    /// </summary>
    private static XlsxWorkbook Charts()
    {
        var book = new XlsxWorkbook { Title = "Combined charts, trendlines, error bars" };
        var sheet = book.AddSheet("Data");
        foreach (var (c, header) in new[] { (1, "Month"), (2, "Pens"), (3, "Paper"), (4, "Share"), (6, "Frequency"), (7, "Gain"), (8, "Error"), (9, "Loss") })
        {
            sheet.Set(1, c, header);
        }

        var months = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat.AbbreviatedMonthNames;
        for (var m = 0; m < 12; m++)
        {
            var row = m + 2;
            var pens = 120 + (m * 9) + ((m % 3) * 14);
            var paper = 80 + (m * 4) + ((m % 4) * 10);
            sheet.Set(row, 1, months[m]);
            sheet.Set(row, 2, pens);
            sheet.Set(row, 3, paper);
            sheet.Set(row, 4, pens / (double)(pens + paper));
            sheet.Format(row, 4, (ref XlsxStyle s) => s.NumberFormat = "0%");

            var frequency = Math.Pow(10, 1 + (m / 4.0));
            sheet.Set(row, 6, frequency);
            sheet.Set(row, 7, 3 + (2.2 * Math.Log10(frequency)) + ((m % 3) * 0.3));
            sheet.Set(row, 8, 0.2 + (m * 0.03));
            sheet.Set(row, 9, 0.5 * m);
        }

        var combined = sheet.AddChart(1, 11, 560, 320);
        combined.Kind = XlsxChartKind.Column;
        combined.Title = "Columns with a line on the right axis";
        combined.YAxis.Title = "Units";
        combined.SecondaryYAxis.Title = "Share of pens";
        combined.SecondaryYAxis.NumberFormat = "0%";
        combined.SecondaryYAxis.Min = 0;
        combined.SecondaryYAxis.Max = 1;
        combined.AddSeries("Pens", 2, 13, 1, 2);
        combined.AddSeries("Paper", 2, 13, 1, 3);
        var share = combined.AddSeries("Share", 2, 13, 1, 4, new XlsxLine("#C62828", 2, "solid"));
        share.Kind = XlsxChartKind.Line;
        share.SecondaryAxis = true;
        share.MarkerSize = 6;

        var mixed = sheet.AddChart(1, 21, 560, 320);
        mixed.Kind = XlsxChartKind.Line;
        mixed.Title = "Area under a trend line";
        mixed.AddSeries("Paper", 2, 13, 1, 3).Kind = XlsxChartKind.Area;
        var pensLine = mixed.AddSeries("Pens", 2, 13, 1, 2);
        pensLine.Trendline = new XlsxTrendline(XlsxTrendlineKind.Linear) { ShowEquation = true, ShowRSquared = true, Name = "Pens trend" };

        var gain = sheet.AddChart(18, 11, 560, 320);
        gain.Title = "Gain over a logarithmic frequency axis";
        gain.XAxis.LogBase = 10;
        gain.XAxis.Title = "Frequency (Hz)";
        gain.YAxis.Title = "Gain (dB)";
        gain.SecondaryYAxis.Title = "Loss (dB)";
        gain.SecondaryYAxis.Reverse = true;
        var measured = gain.AddPoints("Gain", 2, 13, 6, 7);
        measured.ErrorBars = XlsxErrorBars.FromColumns(8);
        measured.Trendline = new XlsxTrendline(XlsxTrendlineKind.Logarithmic) { ShowEquation = true };
        var loss = gain.AddSeries("Loss", 2, 13, 6, 9, new XlsxLine("#70AD47", 1.5, "dash"));
        loss.SecondaryAxis = true;

        var bars = sheet.AddChart(18, 21, 560, 320);
        bars.Kind = XlsxChartKind.Bar;
        bars.Title = "Bars, first month on top, with error bars";
        bars.XAxis.Reverse = true;
        var pensBars = bars.AddSeries("Pens", 2, 7, 1, 2);
        pensBars.ErrorBars = XlsxErrorBars.Percent(5);
        pensBars.Kind = XlsxChartKind.Line;
        var paperBars = bars.AddSeries("Paper", 2, 7, 1, 3);
        paperBars.ErrorBars = XlsxErrorBars.Fixed(10);
        paperBars.ErrorBars.Direction = XlsxErrorDirection.Plus;
        paperBars.ErrorBars.EndCap = false;

        var trends = sheet.AddChart(35, 11, 560, 320);
        trends.Kind = XlsxChartKind.Column;
        trends.Title = "Trendlines and spread";
        var a = trends.AddSeries("Pens", 2, 13, 1, 2);
        a.Trendline = new XlsxTrendline(XlsxTrendlineKind.Polynomial) { Order = 3 };
        a.ErrorBars = XlsxErrorBars.StandardDeviation(1);
        var b = trends.AddSeries("Paper", 2, 13, 1, 3);
        b.Trendline = new XlsxTrendline(XlsxTrendlineKind.MovingAverage) { Period = 3, Line = new XlsxLine("#000000", 2, "solid") };
        b.ErrorBars = XlsxErrorBars.StandardError();
        return book;
    }

    private static SpreadsheetDocument Open(bool keep = false)
    {
        var book = Charts();
        book.Created = XlsxTestKit.FixedTime;
        var path = Path.Combine(XlsxTestKit.NewFolder(), "charts-advanced.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, "charts-advanced.xlsx");
        }

        return SpreadsheetDocument.Open(path, false);
    }

    private static C.PlotArea[] Plots(SpreadsheetDocument document) =>
        document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts
            .Select(p => p.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!).ToArray();

    private static uint[] AxisIds(DocumentFormat.OpenXml.OpenXmlElement group) =>
        group.Elements<C.AxisId>().Select(a => a.Val!.Value).ToArray();

    [Fact]
    public void ColumnsWithALineOnTheRightAxis()
    {
        using (var document = Open(keep: true))
        {
            var plot = Plots(document)[0];
            var columns = plot.GetFirstChild<C.BarChart>()!;
            var line = plot.GetFirstChild<C.LineChart>()!;
            Assert.Equal(2, columns.Elements<C.BarChartSeries>().Count());
            Assert.Single(line.Elements<C.LineChartSeries>());
            Assert.Equal(new uint[] { 1, 2 }, AxisIds(columns));
            Assert.Equal(new uint[] { 3, 4 }, AxisIds(line));

            // Series numbers run across the groups.
            var indexes = plot.Descendants<C.Index>().Where(i => i.Parent is C.BarChartSeries || i.Parent is C.LineChartSeries).Select(i => i.Val!.Value);
            Assert.Equal(new uint[] { 0, 1, 2 }, indexes.OrderBy(i => i).ToArray());

            var categoryAxes = plot.Elements<C.CategoryAxis>().ToList();
            var valueAxes = plot.Elements<C.ValueAxis>().ToList();
            Assert.Equal(2, categoryAxes.Count);
            Assert.False(categoryAxes[0].Delete!.Val!.Value);
            Assert.True(categoryAxes[1].Delete!.Val!.Value);
            Assert.Equal(4u, categoryAxes[1].CrossingAxis!.Val!.Value);
            Assert.Null(categoryAxes[1].Title);
            var right = valueAxes.Single(v => v.AxisId!.Val!.Value == 4);
            Assert.Equal(C.AxisPositionValues.Right, right.AxisPosition!.Val!.Value);
            Assert.Equal(C.CrossesValues.Maximum, right.GetFirstChild<C.Crosses>()!.Val!.Value);
            Assert.Equal("0%", right.NumberingFormat!.FormatCode!.Value);
            Assert.Equal(1, right.Scaling!.MaxAxisValue!.Val!.Value);
            Assert.Equal("Share of pens", right.Title!.InnerText);
        }
    }

    [Fact]
    public void AreaAndLineShareTheAxes()
    {
        using (var document = Open())
        {
            var plot = Plots(document)[1];
            Assert.Equal(new uint[] { 1, 2 }, AxisIds(plot.GetFirstChild<C.AreaChart>()!));
            Assert.Equal(new uint[] { 1, 2 }, AxisIds(plot.GetFirstChild<C.LineChart>()!));
            Assert.Single(plot.Elements<C.ValueAxis>());

            var trend = plot.GetFirstChild<C.LineChart>()!.GetFirstChild<C.LineChartSeries>()!.GetFirstChild<C.Trendline>()!;
            Assert.Equal(C.TrendlineValues.Linear, trend.TrendlineType!.Val!.Value);
            Assert.Equal("Pens trend", trend.TrendlineName!.Text);
            Assert.True(trend.DisplayEquation!.Val!.Value);
            Assert.True(trend.DisplayRSquaredValue!.Val!.Value);
        }
    }

    [Fact]
    public void ScatterWithLogAxisSecondAxisAndCustomErrorBars()
    {
        using (var document = Open())
        {
            var plot = Plots(document)[2];
            var groups = plot.Elements<C.ScatterChart>().ToList();
            Assert.Equal(2, groups.Count);
            Assert.Equal(new uint[] { 3, 4 }, AxisIds(groups[1]));
            var axes = plot.Elements<C.ValueAxis>().ToList();
            Assert.Equal(4, axes.Count);
            Assert.Equal(10, axes[0].Scaling!.LogBase!.Val!.Value);
            Assert.True(axes[2].Delete!.Val!.Value);

            // The hidden X axis carries no title or gridlines; the right axis keeps its labels beside it.
            Assert.Null(axes[2].Title);
            Assert.Null(axes[2].MajorGridlines);
            Assert.Equal(C.TickLabelPositionValues.NextTo, axes[3].TickLabelPosition!.Val!.Value);
            Assert.Equal(C.TickLabelPositionValues.Low, axes[1].TickLabelPosition!.Val!.Value);
            Assert.Equal(10, axes[2].Scaling!.LogBase!.Val!.Value);
            Assert.Equal(C.OrientationValues.MaxMin, axes[3].Scaling!.Orientation!.Val!.Value);
            Assert.Equal(C.AxisPositionValues.Right, axes[3].AxisPosition!.Val!.Value);

            var series = groups[0].GetFirstChild<C.ScatterChartSeries>()!;
            var errors = series.GetFirstChild<C.ErrorBars>()!;
            Assert.Equal(C.ErrorBarDirectionValues.Y, errors.ErrorDirection!.Val!.Value);
            Assert.Equal(C.ErrorValues.Custom, errors.ErrorBarValueType!.Val!.Value);
            Assert.Equal("'Data'!$H$2:$H$13", errors.Plus!.NumberReference!.Formula!.Text);
            Assert.Equal("'Data'!$H$2:$H$13", errors.Minus!.NumberReference!.Formula!.Text);
            Assert.Equal(12u, errors.Plus.NumberReference.NumberingCache!.PointCount!.Val!.Value);
            Assert.Equal(C.TrendlineValues.Logarithmic, series.GetFirstChild<C.Trendline>()!.TrendlineType!.Val!.Value);
        }
    }

    [Fact]
    public void BarChartsKeepTheirKindAndReverseTheirCategories()
    {
        using (var document = Open())
        {
            var plot = Plots(document)[3];
            Assert.Null(plot.GetFirstChild<C.LineChart>());
            var bars = plot.GetFirstChild<C.BarChart>()!.Elements<C.BarChartSeries>().ToList();
            Assert.Equal(2, bars.Count);
            Assert.Equal(C.OrientationValues.MaxMin, plot.GetFirstChild<C.CategoryAxis>()!.Scaling!.Orientation!.Val!.Value);
            Assert.Equal(C.OrientationValues.MinMax, plot.GetFirstChild<C.ValueAxis>()!.Scaling!.Orientation!.Val!.Value);

            // With the categories reversed, the value axis crosses at the last category and stays at the bottom.
            Assert.Equal(C.CrossesValues.Maximum, plot.GetFirstChild<C.ValueAxis>()!.GetFirstChild<C.Crosses>()!.Val!.Value);

            var percent = bars[0].GetFirstChild<C.ErrorBars>()!;
            Assert.Null(percent.ErrorDirection);
            Assert.Equal(C.ErrorValues.Percentage, percent.ErrorBarValueType!.Val!.Value);
            Assert.Equal(5, percent.ErrorBarValue!.Val!.Value);
            Assert.Equal(C.ErrorBarValues.Both, percent.ErrorBarType!.Val!.Value);
            var fixedBars = bars[1].GetFirstChild<C.ErrorBars>()!;
            Assert.Equal(C.ErrorBarValues.Plus, fixedBars.ErrorBarType!.Val!.Value);
            Assert.True(fixedBars.NoEndCap!.Val!.Value);
        }
    }

    [Fact]
    public void PolynomialMovingAverageAndSpread()
    {
        using (var document = Open())
        {
            var series = Plots(document)[4].GetFirstChild<C.BarChart>()!.Elements<C.BarChartSeries>().ToList();
            var poly = series[0].GetFirstChild<C.Trendline>()!;
            Assert.Equal(C.TrendlineValues.Polynomial, poly.TrendlineType!.Val!.Value);
            Assert.Equal((byte)3, poly.PolynomialOrder!.Val!.Value);
            Assert.Equal(C.ErrorValues.StandardDeviation, series[0].GetFirstChild<C.ErrorBars>()!.ErrorBarValueType!.Val!.Value);
            var moving = series[1].GetFirstChild<C.Trendline>()!;
            Assert.Equal(C.TrendlineValues.MovingAverage, moving.TrendlineType!.Val!.Value);
            Assert.Equal(3u, moving.Period!.Val!.Value);
            var stdErr = series[1].GetFirstChild<C.ErrorBars>()!;
            Assert.Equal(C.ErrorValues.StandardError, stdErr.ErrorBarValueType!.Val!.Value);
            Assert.Null(stdErr.ErrorBarValue);
        }
    }

    [Fact]
    public void OutOfRangeSettingsAreClamped()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        for (var r = 1; r <= 5; r++)
        {
            sheet.Set(r, 1, r);
            sheet.Set(r, 2, r * r);
        }

        var chart = sheet.AddChart(1, 4, 300, 200);
        chart.XAxis.LogBase = 1;
        chart.YAxis.LogBase = 5000;
        var s = chart.AddSeries("S", 1, 5, 1, 2);
        s.Trendline = new XlsxTrendline(XlsxTrendlineKind.Polynomial) { Order = 9 };
        var t = chart.AddSeries("T", 1, 5, 1, 2);
        t.Trendline = new XlsxTrendline(XlsxTrendlineKind.MovingAverage) { Period = 0 };
        var path = Path.Combine(XlsxTestKit.NewFolder(), "clamp.xlsx");
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var plot = Plots(document)[0];
            var axes = plot.Elements<C.ValueAxis>().ToList();
            Assert.Equal(2, axes[0].Scaling!.LogBase!.Val!.Value);
            Assert.Equal(1000, axes[1].Scaling!.LogBase!.Val!.Value);
            var trends = plot.Descendants<C.Trendline>().ToList();
            Assert.Equal((byte)6, trends[0].PolynomialOrder!.Val!.Value);
            Assert.Equal(2u, trends[1].Period!.Val!.Value);
        }
    }

    [Fact]
    public void ChartsWithoutExtrasHaveNoSecondAxis()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.Set(1, 1, "a");
        sheet.Set(1, 2, 1.0);
        var chart = sheet.AddChart(1, 4, 300, 200);
        chart.Kind = XlsxChartKind.Column;
        chart.AddSeries("S", 1, 1, 1, 2);
        var path = Path.Combine(XlsxTestKit.NewFolder(), "plain.xlsx");
        book.Save(path);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var plot = Plots(document)[0];
            Assert.Single(plot.Elements<C.CategoryAxis>());
            Assert.Single(plot.Elements<C.ValueAxis>());
            Assert.Empty(plot.Descendants<C.Trendline>());
            Assert.Empty(plot.Descendants<C.ErrorBars>());
        }
    }
}
