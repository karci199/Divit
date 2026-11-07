#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;
using System.Collections.Generic;

/// <summary>
/// An Excel chart embedded in a sheet, with series taken from cell ranges on the sheet: scatter (X – Y), polar,
/// column, bar, line, area, pie or doughnut.
/// </summary>
/// <remarks>
/// This is a native Excel chart: the user can select and edit it, and it follows the cells when their values change.
/// On save the series values are also read from the sheet and written into the chart as a cache, so Excel and
/// OnlyOffice can draw it without recalculating. The look (<see cref="Look"/>) is kept apart from the content, so one
/// look serves every chart.
/// <para>The position is the top-left cell plus a size in pixels; the chart keeps its size when cell sizes change.</para>
/// </remarks>
public sealed class XlsxChart
{
    #region Fields

    private readonly List<XlsxSeries> _series = new List<XlsxSeries>();

    #endregion

    #region Construction

    internal XlsxChart(XlsxSheet sheet, int row, int column, int width, int height)
    {
        Sheet = sheet;
        Row = row;
        Column = column;
        Width = width;
        Height = height;
    }

    #endregion

    #region Properties

    /// <summary>
    /// The sheet the chart sits on; series ranges refer to this sheet.
    /// </summary>
    public XlsxSheet Sheet { get; }

    /// <summary>
    /// Row of the top-left corner (1-based).
    /// </summary>
    public int Row { get; }

    /// <summary>
    /// Column of the top-left corner (1-based).
    /// </summary>
    public int Column { get; }

    /// <summary>
    /// Width in pixels.
    /// </summary>
    public int Width { get; }

    /// <summary>
    /// Height in pixels.
    /// </summary>
    public int Height { get; }

    /// <summary>
    /// Chart kind. Set it before adding series: a series takes its defaults (smoothing) from the kind.
    /// </summary>
    public XlsxChartKind Kind { get; set; } = XlsxChartKind.Scatter;

    /// <summary>
    /// How the series of a column, bar, line or area chart stand to each other: side by side, stacked, or stacked to
    /// 100 %.
    /// </summary>
    public XlsxChartGrouping Grouping { get; set; }

    /// <summary>
    /// Gap between the column or bar groups, in percent of a column's width (0–500).
    /// </summary>
    public int GapWidth { get; set; } = 150;

    /// <summary>
    /// The hole of a doughnut chart, in percent of its size (10–90).
    /// </summary>
    public int HoleSize { get; set; } = 50;

    /// <summary>
    /// Angle of the first slice of a pie or doughnut chart, in degrees clockwise from the top (0–360).
    /// </summary>
    public int FirstSliceAngle { get; set; }

    /// <summary>
    /// What the labels on the points show; none by default. <see cref="XlsxDataLabels.Percent"/> applies to pie and
    /// doughnut charts only.
    /// </summary>
    public XlsxDataLabels DataLabels { get; set; }

    /// <summary>
    /// First title line (large, bold); no title when empty.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Second title line (small), for example "Phi = 0°  ·  Frequency = 1500 MHz".
    /// </summary>
    public string? Subtitle { get; set; }

    /// <summary>
    /// Horizontal axis.
    /// </summary>
    public XlsxAxis XAxis { get; } = new XlsxAxis();

    /// <summary>
    /// Vertical axis.
    /// </summary>
    public XlsxAxis YAxis { get; } = new XlsxAxis();

    /// <summary>
    /// Whether the legend shows below the plot; <c>null</c> shows it when there is more than one series, and always for
    /// pie and doughnut charts (their legend names the slices).
    /// </summary>
    public bool? ShowLegend { get; set; }

    /// <summary>
    /// Position of the plot area (inside the axes) within the chart, as fractions (0–1): left, top, width, height.
    /// Excel places it itself when empty. Useful when an image with the same axes sits under the chart (a heat map)
    /// and both must line up pixel for pixel.
    /// </summary>
    public (double X, double Y, double Width, double Height)? PlotLayout { get; set; }

    /// <summary>
    /// The look; defaults to <see cref="XlsxChartLook.Default"/>.
    /// </summary>
    public XlsxChartLook Look { get; set; } = XlsxChartLook.Default;

    /// <summary>
    /// The series in the order they were added.
    /// </summary>
    public IReadOnlyList<XlsxSeries> Series => _series;

    #endregion

    #region Public methods

    /// <summary>
    /// Adds a series from two columns of the sheet; X and Y share the same row range. In a column, bar, line, area, pie
    /// or doughnut chart the X column holds the categories (text, numbers or dates) and the Y column the values.
    /// </summary>
    /// <param name="name">Name shown in the legend.</param>
    /// <param name="firstRow">First data row.</param>
    /// <param name="lastRow">Last data row.</param>
    /// <param name="xColumn">Column of the X values or categories.</param>
    /// <param name="yColumn">Column of the Y values.</param>
    /// <param name="line">Line style, whose colour also fills columns, bars and areas; <c>null</c> uses the look: its
    /// series line for scatter and polar charts, the next palette colour for the others.</param>
    /// <returns>The new series.</returns>
    /// <exception cref="ArgumentException">The row range is empty.</exception>
    public XlsxSeries AddSeries(string name, int firstRow, int lastRow, int xColumn, int yColumn, XlsxLine? line = null)
    {
        if (lastRow < firstRow)
        {
            throw new ArgumentException(Messages.Format(Strings.SeriesRangeEmpty, firstRow, lastRow));
        }

        var series = new XlsxSeries(name, firstRow, lastRow, xColumn, yColumn, line ?? Look.SeriesLine)
        {
            OwnColor = line != null,
            Smooth = Kind == XlsxChartKind.Scatter,
        };
        _series.Add(series);
        return series;
    }

    /// <summary>
    /// Adds one series per column of a table on the sheet: the header row names the series, the category column
    /// labels the points. The usual way to chart a table of figures (months down, products across).
    /// </summary>
    /// <param name="headerRow">Row of the series names.</param>
    /// <param name="lastRow">Last data row; data starts below the header row.</param>
    /// <param name="categoryColumn">Column of the categories.</param>
    /// <param name="firstValueColumn">First value column.</param>
    /// <param name="lastValueColumn">Last value column.</param>
    /// <returns>The new series, in column order.</returns>
    /// <exception cref="ArgumentException">There is no data row or no value column.</exception>
    public IReadOnlyList<XlsxSeries> AddSeriesFromColumns(int headerRow, int lastRow, int categoryColumn, int firstValueColumn, int lastValueColumn)
    {
        if (lastValueColumn < firstValueColumn)
        {
            throw new ArgumentException(Messages.Get(Strings.RangeReversed));
        }

        var added = new List<XlsxSeries>();
        for (var c = firstValueColumn; c <= lastValueColumn; c++)
        {
            added.Add(AddSeries(Sheet.TextAt(headerRow, c) ?? XlsxWorkbook.ColumnName(c), headerRow + 1, lastRow, categoryColumn, c));
        }

        return added;
    }

    /// <summary>
    /// Adds a series drawn as markers only, without a line; for example points outside a limit, where only those rows
    /// of the Y column hold a value and the others are empty (empty cells are not drawn).
    /// </summary>
    /// <param name="name">Name shown in the legend.</param>
    /// <param name="firstRow">First data row.</param>
    /// <param name="lastRow">Last data row.</param>
    /// <param name="xColumn">Column of the X values.</param>
    /// <param name="yColumn">Column of the Y values.</param>
    /// <param name="color">Marker colour; the look's accent colour when empty.</param>
    /// <returns>The new series.</returns>
    public XlsxSeries AddPoints(string name, int firstRow, int lastRow, int xColumn, int yColumn, string? color = null)
    {
        var series = AddSeries(name, firstRow, lastRow, xColumn, yColumn, Look.AccentLine);
        series.ShowLine = false;
        series.MarkerColor = color ?? Look.AccentLine.Color;
        series.MarkerSize = 7;
        return series;
    }

    #endregion
}

/// <summary>
/// Chart kind.
/// </summary>
public enum XlsxChartKind
{
    /// <summary>
    /// Scatter: X and Y value axes, smooth line.
    /// </summary>
    Scatter,

    /// <summary>
    /// Polar (Excel's radar chart): the X column holds angles, which must be evenly spaced; 0° is at the top and angles
    /// run clockwise. Y is the radius and may be negative (dB); the minimum of <see cref="XlsxChart.YAxis"/> becomes
    /// the centre. A constant series draws a circle.
    /// </summary>
    Polar,

    /// <summary>
    /// Vertical columns over categories.
    /// </summary>
    Column,

    /// <summary>
    /// Horizontal bars; Excel lists the first category at the bottom.
    /// </summary>
    Bar,

    /// <summary>
    /// Lines over categories, evenly spaced (for uneven X values use <see cref="Scatter"/>).
    /// </summary>
    Line,

    /// <summary>
    /// Filled areas over categories.
    /// </summary>
    Area,

    /// <summary>
    /// A pie of the first series; the categories name the slices.
    /// </summary>
    Pie,

    /// <summary>
    /// A ring per series, with a hole in the middle (<see cref="XlsxChart.HoleSize"/>).
    /// </summary>
    Doughnut,
}

/// <summary>
/// How the series of a column, bar, line or area chart stand to each other.
/// </summary>
public enum XlsxChartGrouping
{
    /// <summary>
    /// Side by side (columns, bars) or each on its own (lines, areas).
    /// </summary>
    Standard,

    /// <summary>
    /// On top of each other: the top shows the sum.
    /// </summary>
    Stacked,

    /// <summary>
    /// On top of each other, scaled to 100 %: the shares of the sum.
    /// </summary>
    PercentStacked,
}

/// <summary>
/// What the labels on the points of a chart show; values combine.
/// </summary>
[Flags]
public enum XlsxDataLabels
{
    /// <summary>
    /// No labels.
    /// </summary>
    None = 0,

    /// <summary>
    /// The value.
    /// </summary>
    Value = 1,

    /// <summary>
    /// The share of the whole (pie and doughnut charts).
    /// </summary>
    Percent = 2,

    /// <summary>
    /// The category name.
    /// </summary>
    Category = 4,

    /// <summary>
    /// The series name.
    /// </summary>
    SeriesName = 8,
}

/// <summary>
/// An axis: title, range, step and number format. An empty field is left to Excel, which picks a value itself.
/// </summary>
public sealed class XlsxAxis
{
    /// <summary>
    /// Axis title; no title when empty.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Lowest value on the axis; Excel picks it when empty.
    /// </summary>
    public double? Min { get; set; }

    /// <summary>
    /// Highest value on the axis; Excel picks it when empty.
    /// </summary>
    public double? Max { get; set; }

    /// <summary>
    /// Major unit: the step of the gridlines and labels.
    /// </summary>
    public double? MajorUnit { get; set; }

    /// <summary>
    /// Number format of the labels, for example <c>0</c>; the source cells' format when empty.
    /// </summary>
    public string? NumberFormat { get; set; }

    /// <summary>
    /// Whether gridlines are shown; <c>null</c> shows them on value axes and leaves them off on category axes, as
    /// Excel does.
    /// </summary>
    public bool? Gridlines { get; set; }
}

/// <summary>
/// A series of a chart: the X and Y ranges on the sheet, a name and a line style.
/// </summary>
public sealed class XlsxSeries
{
    internal XlsxSeries(string name, int firstRow, int lastRow, int xColumn, int yColumn, XlsxLine line)
    {
        Name = name;
        FirstRow = firstRow;
        LastRow = lastRow;
        XColumn = xColumn;
        YColumn = yColumn;
        Line = line;
    }

    /// <summary>
    /// Name shown in the legend.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// First data row (1-based).
    /// </summary>
    public int FirstRow { get; }

    /// <summary>
    /// Last data row (1-based).
    /// </summary>
    public int LastRow { get; }

    /// <summary>
    /// Column of the X values (1-based).
    /// </summary>
    public int XColumn { get; }

    /// <summary>
    /// Column of the Y values (1-based).
    /// </summary>
    public int YColumn { get; }

    /// <summary>
    /// Line style: colour, width, dash.
    /// </summary>
    public XlsxLine Line { get; set; }

    /// <summary>
    /// Whether the points are joined by a smooth curve rather than straight segments.
    /// </summary>
    public bool Smooth { get; set; } = true;

    /// <summary>
    /// Whether the line is drawn; when off only the markers are shown.
    /// </summary>
    public bool ShowLine { get; set; } = true;

    /// <summary>
    /// Marker (circle) size, 2–72; 0 means no marker.
    /// </summary>
    public int MarkerSize { get; set; }

    /// <summary>
    /// Marker colour; the line colour when empty.
    /// </summary>
    public string? MarkerColor { get; set; }

    /// <summary>
    /// Whether the series appears in the legend. When one line is split into several series only the first one needs
    /// to be listed.
    /// </summary>
    public bool ShowInLegend { get; set; } = true;

    /// <summary>
    /// How far the slices of a pie or doughnut chart stand out from the centre, in percent of the radius (0–400).
    /// </summary>
    public int Explosion { get; set; }

    /// <summary>
    /// Whether the line style was given; otherwise column-like charts colour the series from the look's palette.
    /// </summary>
    internal bool OwnColor { get; set; }
}

/// <summary>
/// Line style: colour, width and dash.
/// </summary>
public readonly struct XlsxLine
{
    /// <summary>
    /// Creates a line style.
    /// </summary>
    /// <param name="color">Colour as <c>#RRGGBB</c>.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="dash">Excel dash name: <c>solid</c>, <c>dash</c>, <c>sysDash</c>, <c>dot</c> …</param>
    public XlsxLine(string color, double width, string dash)
    {
        Color = color;
        Width = width;
        Dash = dash;
    }

    /// <summary>
    /// Colour as <c>#RRGGBB</c>.
    /// </summary>
    public string Color { get; }

    /// <summary>
    /// Width in points.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// Excel dash name (<c>solid</c>, <c>dash</c> …).
    /// </summary>
    public string Dash { get; }
}

/// <summary>
/// The look of a chart, used like a template: font sizes and colours, gridlines, border and the default line styles.
/// </summary>
/// <remarks>
/// Kept apart from the content (data, title text, axis ranges): defined once and shared by every chart, so changing the
/// look means changing only this object.
/// </remarks>
public sealed class XlsxChartLook
{
    /// <summary>
    /// The default look: dark blue bold title, grey subtitle, light grey gridlines, a dark blue series line and a red
    /// dashed accent line.
    /// </summary>
    public static readonly XlsxChartLook Default = new XlsxChartLook();

    /// <summary>
    /// Font name; the workbook default when empty.
    /// </summary>
    public string? FontName { get; set; }

    /// <summary>
    /// Title font size in points.
    /// </summary>
    public double TitleSize { get; set; } = 14;

    /// <summary>
    /// Title colour.
    /// </summary>
    public string TitleColor { get; set; } = "#1F3864";

    /// <summary>
    /// Subtitle font size in points.
    /// </summary>
    public double SubtitleSize { get; set; } = 10;

    /// <summary>
    /// Subtitle colour.
    /// </summary>
    public string SubtitleColor { get; set; } = "#404040";

    /// <summary>
    /// Font size of the axis labels and the legend, in points.
    /// </summary>
    public double TextSize { get; set; } = 9;

    /// <summary>
    /// Colour of the axis labels and the legend.
    /// </summary>
    public string TextColor { get; set; } = "#404040";

    /// <summary>
    /// Axis title font size in points.
    /// </summary>
    public double AxisTitleSize { get; set; } = 9;

    /// <summary>
    /// Gridline colour.
    /// </summary>
    public string GridColor { get; set; } = "#E0E0E0";

    /// <summary>
    /// Axis line colour.
    /// </summary>
    public string AxisColor { get; set; } = "#A6A6A6";

    /// <summary>
    /// Border of the chart area; no border when empty.
    /// </summary>
    public string? BorderColor { get; set; } = "#BFBFBF";

    /// <summary>
    /// Default line of a series.
    /// </summary>
    public XlsxLine SeriesLine { get; set; } = new XlsxLine("#1F4E9A", 2.5, "solid");

    /// <summary>
    /// Accent line, for example a limit; also the default colour of <see cref="XlsxChart.AddPoints"/>.
    /// </summary>
    public XlsxLine AccentLine { get; set; } = new XlsxLine("#C62828", 1.75, "dash");

    /// <summary>
    /// Muted line for secondary references: grey and dotted.
    /// </summary>
    public XlsxLine MutedLine { get; set; } = new XlsxLine("#8C8C8C", 1.5, "sysDot");

    /// <summary>
    /// Colours of the series of column, bar, line and area charts, and of the slices of pie and doughnut charts, used
    /// in turn; the accents of the default Office theme.
    /// </summary>
    public IReadOnlyList<string> Palette { get; set; } = new[] { "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47" };

    /// <summary>
    /// Width of the lines of a line chart, in points.
    /// </summary>
    public double LineWidth { get; set; } = 2.25;
}
