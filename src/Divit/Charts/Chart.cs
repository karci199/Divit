#nullable enable
namespace Divit.Charts;

using Divit.Xlsx;
using System;
using System.Collections.Generic;

/// <summary>
/// The kind of a chart.
/// </summary>
public enum ChartKind
{
    /// <summary>
    /// Lines through x-y points, for measurements against a number such as frequency.
    /// </summary>
    Scatter,

    /// <summary>
    /// Lines over categories (months, test steps).
    /// </summary>
    Line,

    /// <summary>
    /// Vertical bars per category.
    /// </summary>
    Column,

    /// <summary>
    /// Horizontal bars per category.
    /// </summary>
    Bar,

    /// <summary>
    /// Filled areas over categories.
    /// </summary>
    Area,

    /// <summary>
    /// Slices of a whole.
    /// </summary>
    Pie,

    /// <summary>
    /// A pie with a hole.
    /// </summary>
    Doughnut,
}

/// <summary>
/// A chart from values, for Word documents (<see cref="Divit.Docx.DocxParagraph.AddChart"/>) and PDF pages
/// (<see cref="Divit.Pdf.PdfPage.DrawChart(Chart, double, double, double, double, string)"/>): series, limit lines,
/// titles and axes. In Word it is a real chart with its data in an embedded workbook, which Word can edit; in PDF it
/// is drawn as lines and shapes.
/// </summary>
public sealed class Chart
{
    private readonly List<ChartSeries> _series = new List<ChartSeries>();

    /// <summary>
    /// Creates an empty chart.
    /// </summary>
    /// <param name="kind">The kind of chart.</param>
    public Chart(ChartKind kind = ChartKind.Scatter)
    {
        Kind = kind;
    }

    /// <summary>
    /// The kind of chart.
    /// </summary>
    public ChartKind Kind { get; }

    /// <summary>
    /// Title above the chart; none when empty.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Smaller line under the title; none when empty.
    /// </summary>
    public string? Subtitle { get; set; }

    /// <summary>
    /// The horizontal axis (the categories of column, line and area charts); its title, range and number format.
    /// </summary>
    public XlsxAxis XAxis { get; } = new XlsxAxis();

    /// <summary>
    /// The vertical value axis.
    /// </summary>
    public XlsxAxis YAxis { get; } = new XlsxAxis();

    /// <summary>
    /// The value axis on the right, for series that use it.
    /// </summary>
    public XlsxAxis SecondaryYAxis { get; } = new XlsxAxis();

    /// <summary>
    /// Shows the legend; by default when there is more than one series, and for pies.
    /// </summary>
    public bool? ShowLegend { get; set; }

    /// <summary>
    /// Fonts, sizes and colours, shared with Excel charts.
    /// </summary>
    public XlsxChartLook Look { get; set; } = XlsxChartLook.Default;

    /// <summary>
    /// The series in order.
    /// </summary>
    public IReadOnlyList<ChartSeries> Series => _series;

    /// <summary>
    /// Adds a series of x-y points, for scatter charts.
    /// </summary>
    /// <param name="name">Name in the legend.</param>
    /// <param name="x">The x values.</param>
    /// <param name="y">The y values, as many as x.</param>
    /// <returns>The series, to set its colour and line.</returns>
    public ChartSeries AddSeries(string name, IReadOnlyList<double> x, IReadOnlyList<double> y) =>
        Add(new ChartSeries(name, x ?? throw new ArgumentNullException(nameof(x)), null, y ?? throw new ArgumentNullException(nameof(y))));

    /// <summary>
    /// Adds a series of values per category, for line, column, bar, area and pie charts. The first series gives the
    /// categories.
    /// </summary>
    /// <param name="name">Name in the legend.</param>
    /// <param name="categories">The categories.</param>
    /// <param name="values">The values, one per category.</param>
    /// <returns>The series, to set its colour.</returns>
    public ChartSeries AddSeries(string name, IReadOnlyList<string> categories, IReadOnlyList<double> values) =>
        Add(new ChartSeries(name, null, categories ?? throw new ArgumentNullException(nameof(categories)), values ?? throw new ArgumentNullException(nameof(values))));

    /// <summary>
    /// Adds points without a line, such as measured maxima on a curve.
    /// </summary>
    /// <param name="name">Name in the legend.</param>
    /// <param name="x">The x values.</param>
    /// <param name="y">The y values.</param>
    /// <param name="color">Colour; the look's accent when empty.</param>
    /// <returns>The series.</returns>
    public ChartSeries AddPoints(string name, IReadOnlyList<double> x, IReadOnlyList<double> y, string? color = null)
    {
        var series = AddSeries(name, x, y);
        series.Color = color ?? Look.AccentLine.Color;
        series.ShowLine = false;
        series.MarkerSize = 7;
        return series;
    }

    /// <summary>
    /// Adds a limit: a dashed horizontal line across the chart at a value, such as an upper limit of 2.0.
    /// </summary>
    /// <param name="name">Name in the legend.</param>
    /// <param name="value">The value.</param>
    /// <param name="color">Colour; the look's accent when empty.</param>
    /// <returns>The series that draws it.</returns>
    public ChartSeries AddLimit(string name, double value, string? color = null)
    {
        var series = new ChartSeries(name, null, null, new[] { value, value }) { Limit = true, Color = color ?? Look.AccentLine.Color, Width = Look.AccentLine.Width, Dash = "dash" };
        return Add(series);
    }

    private ChartSeries Add(ChartSeries series)
    {
        _series.Add(series);
        return series;
    }

    /// <summary>
    /// The chart as an Excel chart over a sheet of its values: the form Word embeds and the PDF painter draws.
    /// Series of points take two columns each; series of categories share the first series' categories; limits
    /// span the x range of the points, or every category.
    /// </summary>
    internal XlsxChart Build(int width = 640, int height = 360) => Build(width, height, out _);

    /// <summary>
    /// The chart as an Excel chart, and the workbook that holds its values, to embed.
    /// </summary>
    internal XlsxChart Build(int width, int height, out XlsxWorkbook book)
    {
        book = new XlsxWorkbook { Created = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc) };
        var sheet = book.AddSheet("Data");
        var scatter = Kind == ChartKind.Scatter;
        var kind = Kind switch
        {
            ChartKind.Line => XlsxChartKind.Line,
            ChartKind.Column => XlsxChartKind.Column,
            ChartKind.Bar => XlsxChartKind.Bar,
            ChartKind.Area => XlsxChartKind.Area,
            ChartKind.Pie => XlsxChartKind.Pie,
            ChartKind.Doughnut => XlsxChartKind.Doughnut,
            _ => XlsxChartKind.Scatter,
        };

        // The categories, or the x range a limit spans.
        IReadOnlyList<string>? categories = null;
        double low = double.MaxValue, high = double.MinValue;
        foreach (var s in _series)
        {
            categories ??= s.Categories;
            if (s.X != null)
            {
                foreach (var x in s.X)
                {
                    low = Math.Min(low, x);
                    high = Math.Max(high, x);
                }
            }
        }

        if (low > high)
        {
            low = 0;
            high = 1;
        }

        var column = 1;
        var rows = 1;
        if (!scatter && categories != null)
        {
            sheet.Set(1, 1, XAxis.Title ?? "Category");
            for (var i = 0; i < categories.Count; i++)
            {
                sheet.Set(i + 2, 1, categories[i]);
            }

            rows = categories.Count + 1;
            column = 2;
        }

        var chart = sheet.AddChart(1, 1, width, height);
        chart.Kind = kind;
        chart.Title = Title;
        chart.Subtitle = Subtitle;
        chart.ShowLegend = ShowLegend;
        chart.Look = Look;
        Copy(XAxis, chart.XAxis);
        Copy(YAxis, chart.YAxis);
        Copy(SecondaryYAxis, chart.SecondaryYAxis);
        var look = Look;
        for (var i = 0; i < _series.Count; i++)
        {
            var s = _series[i];
            var color = s.Color ?? (i == 0 || look.Palette.Count == 0 ? look.SeriesLine.Color : look.Palette[i % look.Palette.Count]);
            var line = new XlsxLine(color, s.Width ?? (scatter ? look.SeriesLine.Width : look.LineWidth), s.Dash ?? "solid");
            XlsxSeries added;
            if (scatter)
            {
                var xs = s.Limit ? new[] { low, high } : (IReadOnlyList<double>)(s.X ?? Index(s.Values.Count));
                sheet.Set(1, column, s.Name + " x");
                sheet.Set(1, column + 1, s.Name);
                var count = Math.Min(xs.Count, s.Values.Count);
                for (var r = 0; r < count; r++)
                {
                    sheet.Set(r + 2, column, xs[r]);
                    sheet.Set(r + 2, column + 1, s.Values[r]);
                }

                added = chart.AddSeries(s.Name, 2, Math.Max(2, count + 1), column, column + 1, line);
                added.Smooth = false;
                column += 2;
            }
            else
            {
                sheet.Set(1, column, s.Name);
                var count = Math.Max(1, rows - 1);
                for (var r = 0; r < count; r++)
                {
                    var value = s.Limit ? s.Values[0] : r < s.Values.Count ? s.Values[r] : double.NaN;
                    if (!double.IsNaN(value))
                    {
                        sheet.Set(r + 2, column, value);
                    }
                }

                added = chart.AddSeries(s.Name, 2, count + 1, 1, column, s.Color != null || s.Limit ? line : (XlsxLine?)null);
                if (s.Limit && kind != XlsxChartKind.Line)
                {
                    // A limit over bars or areas is a line on top of them.
                    added.Kind = XlsxChartKind.Line;
                }

                column++;
            }

            added.ShowLine = s.ShowLine;
            added.MarkerSize = s.MarkerSize;
            added.MarkerColor = s.MarkerSize > 0 ? color : null;
            added.SecondaryAxis = s.SecondaryAxis;
        }

        return chart;
    }

    private static double[] Index(int count)
    {
        var x = new double[count];
        for (var i = 0; i < count; i++)
        {
            x[i] = i + 1;
        }

        return x;
    }

    private static void Copy(XlsxAxis from, XlsxAxis to)
    {
        to.Title = from.Title;
        to.Min = from.Min;
        to.Max = from.Max;
        to.MajorUnit = from.MajorUnit;
        to.NumberFormat = from.NumberFormat;
        to.Gridlines = from.Gridlines;
        to.LogBase = from.LogBase;
        to.Reverse = from.Reverse;
    }
}

/// <summary>
/// A series of a <see cref="Chart"/>: its values, and how it is drawn.
/// </summary>
public sealed class ChartSeries
{
    internal ChartSeries(string name, IReadOnlyList<double>? x, IReadOnlyList<string>? categories, IReadOnlyList<double> values)
    {
        Name = name ?? string.Empty;
        X = x;
        Categories = categories;
        Values = values;
    }

    /// <summary>
    /// Name in the legend.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Colour as <c>#RRGGBB</c>; the look's when empty.
    /// </summary>
    public string? Color { get; set; }

    /// <summary>
    /// Line width in points; the look's when empty.
    /// </summary>
    public double? Width { get; set; }

    /// <summary>
    /// Line pattern: <c>solid</c>, <c>dash</c>, <c>sysDot</c>, <c>sysDash</c>, <c>dashDot</c>, <c>lgDash</c>.
    /// </summary>
    public string? Dash { get; set; }

    /// <summary>
    /// Draws the line; off for points only.
    /// </summary>
    public bool ShowLine { get; set; } = true;

    /// <summary>
    /// Size of a circle at each point in points; none when 0.
    /// </summary>
    public int MarkerSize { get; set; }

    /// <summary>
    /// Plots the series against the value axis on the right.
    /// </summary>
    public bool SecondaryAxis { get; set; }

    internal IReadOnlyList<double>? X { get; }

    internal IReadOnlyList<string>? Categories { get; }

    internal IReadOnlyList<double> Values { get; }

    /// <summary>
    /// Whether the series is a limit line across the chart.
    /// </summary>
    internal bool Limit { get; set; }
}
