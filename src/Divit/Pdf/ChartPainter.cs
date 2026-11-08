#nullable enable
namespace Divit.Pdf;

using Divit.Xlsx;
using System;
using System.Collections.Generic;
using System.Globalization;

/// <summary>
/// Draws an Excel chart on a PDF page as lines, shapes and text, close to how Excel shows it: titles, axes with
/// round tick steps, gridlines, the series (scatter, line, column, bar, area, pie, doughnut, combined), a second
/// value axis, and the legend under the plot. Values come from the cells the series point at.
/// </summary>
internal static class ChartPainter
{
    private const double Pad = 8;

    public static void Draw(PdfPage page, XlsxChart chart, double x, double y, double width, double height, string defaultFont)
    {
        var look = chart.Look;
        var family = look.FontName ?? defaultFont;
        var regular = Font(page, family, false);
        var bold = Font(page, family, true);
        if (look.BorderColor != null)
        {
            page.StrokeRectangle(x, y, width, height, look.BorderColor, 0.75);
        }

        double left = x + Pad, right = x + width - Pad, top = y + Pad, bottom = y + height - Pad;
        if (!string.IsNullOrEmpty(chart.Title))
        {
            top += look.TitleSize;
            page.DrawText(chart.Title!, (left + right) / 2, top, bold, look.TitleSize, look.TitleColor, PdfAlign.Center);
            top += look.TitleSize * 0.45;
        }

        if (!string.IsNullOrEmpty(chart.Subtitle))
        {
            top += look.SubtitleSize;
            page.DrawText(chart.Subtitle!, (left + right) / 2, top, regular, look.SubtitleSize, look.SubtitleColor, PdfAlign.Center);
            top += look.SubtitleSize * 0.45;
        }

        top += 6;
        var series = Read(chart);
        var pie = chart.Kind == XlsxChartKind.Pie || chart.Kind == XlsxChartKind.Doughnut;
        if ((chart.ShowLegend ?? (pie || chart.Series.Count > 1)) && series.Count > 0)
        {
            bottom = Legend(page, chart, series, pie, regular, left, right, bottom);
        }

        if (series.Count == 0 || right - left < 20 || bottom - top < 20)
        {
            return;
        }

        if (pie)
        {
            Pie(page, chart, series[0], regular, left, top, right, bottom);
        }
        else
        {
            Plot(page, chart, series, regular, bold, left, top, right, bottom);
        }
    }

    private static PdfFont Font(PdfPage page, string family, bool bold)
    {
        foreach (var name in new[] { family, "Calibri", "Arial", "Liberation Sans", "DejaVu Sans" })
        {
            try
            {
                return page.Document.AddFont(name, bold);
            }
            catch (ArgumentException)
            {
            }
            catch (NotSupportedException)
            {
            }
        }

        throw new InvalidOperationException(Divit.Properties.Messages.Format(Divit.Properties.Strings.FontNotFound, family));
    }

    // ------------------------------------------------------------------------------------------------- data

    /// <summary>
    /// The values of a series: x-y points, or a label and a value per category; its colour and line.
    /// </summary>
    private sealed class Data
    {
        public XlsxSeries Source = null!;
        public XlsxChartKind Kind;
        public string Color = "#000000";
        public double Width;
        public PdfDash Dash;
        public List<string> Labels { get; } = new List<string>();
        public List<double?> X { get; } = new List<double?>();
        public List<double?> Y { get; } = new List<double?>();
    }

    private static List<Data> Read(XlsxChart chart)
    {
        var look = chart.Look;
        var sheet = chart.Sheet;
        var list = new List<Data>();
        for (var i = 0; i < chart.Series.Count; i++)
        {
            var s = chart.Series[i];
            var kind = s.Kind ?? chart.Kind;
            var category = kind != XlsxChartKind.Scatter && kind != XlsxChartKind.Polar;
            var color = category && !s.OwnColor ? (look.Palette.Count == 0 ? look.SeriesLine.Color : look.Palette[i % look.Palette.Count]) : s.Line.Color;
            var width = category && !s.OwnColor ? look.LineWidth : s.Line.Width;
            var data = new Data { Source = s, Kind = kind, Color = "#" + Divit.Packaging.OfficeText.Rgb(color), Width = width, Dash = Dash(category && !s.OwnColor ? "solid" : s.Line.Dash) };
            for (var r = s.FirstRow; r <= s.LastRow; r++)
            {
                var x = sheet.NumberAt(r, s.XColumn);
                data.X.Add(x);
                data.Labels.Add(sheet.TextAt(r, s.XColumn) ?? (x is double v ? Number(v, sheet.FormatAt(r, s.XColumn), 0) : string.Empty));
                data.Y.Add(sheet.NumberAt(r, s.YColumn));
            }

            list.Add(data);
        }

        return list;
    }

    private static PdfDash Dash(string? dash) =>
        dash == null || dash == "solid" ? PdfDash.Solid : dash.IndexOf("Dot", StringComparison.OrdinalIgnoreCase) >= 0 && dash.IndexOf("Dash", StringComparison.OrdinalIgnoreCase) < 0 ? PdfDash.Dotted : PdfDash.Dashed;

    // ------------------------------------------------------------------------------------------------- scales

    /// <summary>
    /// A value axis: its range, tick step, and whether it is logarithmic or reversed.
    /// </summary>
    private sealed class Scale
    {
        public double Low;
        public double High;
        public double Step;
        public double? Log;
        public bool Reverse;
        public string? Format;

        /// <summary>
        /// Where a value lies between the ends, 0 at the low end.
        /// </summary>
        public double Fraction(double value)
        {
            double f;
            if (Log is double b)
            {
                f = (Math.Log(Math.Max(value, Low * 1e-9)) - Math.Log(Low)) / (Math.Log(High) - Math.Log(Low));
            }
            else
            {
                f = (value - Low) / (High - Low);
            }

            return Reverse ? 1 - f : f;
        }

        public IEnumerable<double> Ticks()
        {
            if (Log is double b)
            {
                for (var v = Low; v <= High * 1.0001; v *= b)
                {
                    yield return v;
                }

                yield break;
            }

            var count = (int)Math.Round((High - Low) / Step);
            for (var i = 0; i <= count && i <= 1000; i++)
            {
                yield return Low + (i * Step);
            }
        }
    }

    /// <summary>
    /// A range with round steps (1, 2, 2.5, 5 times a power of ten) around the values, or the axis's own.
    /// </summary>
    private static Scale Range(XlsxAxis axis, double min, double max, bool zero)
    {
        var scale = new Scale { Reverse = axis.Reverse, Format = axis.NumberFormat };
        if (axis.LogBase is double b && b > 1)
        {
            var low = axis.Min ?? Math.Pow(b, Math.Floor(Math.Log(Math.Max(min, 1e-12)) / Math.Log(b)));
            var high = axis.Max ?? Math.Pow(b, Math.Ceiling(Math.Log(Math.Max(max, low * b)) / Math.Log(b)));
            scale.Low = low;
            scale.High = high <= low ? low * b : high;
            scale.Log = b;
            scale.Step = b;
            return scale;
        }

        if (zero)
        {
            min = Math.Min(min, 0);
            max = Math.Max(max, 0);
        }

        if (axis.Min is double fixedLow) { min = fixedLow; }
        if (axis.Max is double fixedHigh) { max = fixedHigh; }
        if (max <= min)
        {
            max = min + (min == 0 ? 1 : Math.Abs(min) * 0.1);
        }

        var step = axis.MajorUnit ?? Nice((max - min) / 5);
        scale.Step = step;
        scale.Low = axis.Min ?? Math.Floor((min / step) + 1e-9) * step;
        scale.High = axis.Max ?? Math.Ceiling((max / step) - 1e-9) * step;
        if (scale.High <= scale.Low)
        {
            scale.High = scale.Low + step;
        }

        return scale;
    }

    private static double Nice(double rough)
    {
        if (rough <= 0 || double.IsNaN(rough) || double.IsInfinity(rough))
        {
            return 1;
        }

        var magnitude = Math.Pow(10, Math.Floor(Math.Log10(rough)));
        foreach (var factor in new[] { 1, 2, 2.5, 5, 10 })
        {
            if (factor * magnitude >= rough * 0.999)
            {
                return factor * magnitude;
            }
        }

        return 10 * magnitude;
    }

    /// <summary>
    /// A number as an axis label: the decimals of the format (<c>0.00</c>, <c>0%</c>), else as many as the step
    /// needs; a minus sign instead of a hyphen.
    /// </summary>
    private static string Number(double value, string? format, double step)
    {
        int decimals;
        var percent = format != null && format.IndexOf('%') >= 0;
        if (format != null && format != "General")
        {
            var point = format.IndexOf('.');
            decimals = 0;
            if (point >= 0)
            {
                for (var i = point + 1; i < format.Length && (format[i] == '0' || format[i] == '#'); i++)
                {
                    decimals++;
                }
            }
        }
        else if (step > 0 && step < 1)
        {
            decimals = Math.Min(6, (int)Math.Ceiling(-Math.Log10(step) - 1e-9));
        }
        else if (step == 0)
        {
            return value.ToString("0.######", CultureInfo.InvariantCulture).Replace("-", "−");
        }
        else
        {
            decimals = 0;
        }

        var shown = percent ? value * 100 : value;
        if (Math.Abs(shown) < Math.Pow(10, -decimals) / 2)
        {
            shown = 0;
        }

        return (shown.ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture) + (percent ? "%" : string.Empty)).Replace("-", "−");
    }

    // ------------------------------------------------------------------------------------------------- plots

    private static void Plot(PdfPage page, XlsxChart chart, List<Data> series, PdfFont regular, PdfFont bold, double left, double top, double right, double bottom)
    {
        var look = chart.Look;
        var text = look.TextSize;
        var scatter = chart.Kind == XlsxChartKind.Scatter || chart.Kind == XlsxChartKind.Polar;
        var horizontal = chart.Kind == XlsxChartKind.Bar;
        var stacked = chart.Grouping != XlsxChartGrouping.Standard;
        var percent = chart.Grouping == XlsxChartGrouping.PercentStacked;
        var categories = scatter ? 0 : Count(series);
        var labels = scatter ? new List<string>() : series[0].Labels;

        // The value ranges: x and y of points, or the values (summed when stacked) of each axis.
        double xMin = double.MaxValue, xMax = double.MinValue;
        double[] yMin = { double.MaxValue, double.MaxValue }, yMax = { double.MinValue, double.MinValue };
        var hasSecondary = false;
        foreach (var s in series)
        {
            var axis = s.Source.SecondaryAxis ? 1 : 0;
            hasSecondary |= axis == 1;
            for (var i = 0; i < s.Y.Count; i++)
            {
                if (s.Y[i] is double v)
                {
                    yMin[axis] = Math.Min(yMin[axis], v);
                    yMax[axis] = Math.Max(yMax[axis], v);
                }

                if (scatter && s.X[i] is double xv)
                {
                    xMin = Math.Min(xMin, xv);
                    xMax = Math.Max(xMax, xv);
                }
            }
        }

        if (stacked)
        {
            for (var c = 0; c < categories; c++)
            {
                double up = 0, down = 0;
                foreach (var s in series)
                {
                    if (s.Kind != chart.Kind || c >= s.Y.Count || !(s.Y[c] is double v)) { continue; }
                    if (v >= 0) { up += v; } else { down += v; }
                }

                yMin[0] = Math.Min(yMin[0], percent ? (down < 0 ? -1 : 0) : down);
                yMax[0] = Math.Max(yMax[0], percent ? (up > 0 ? 1 : 0) : up);
            }
        }

        var columns = chart.Kind == XlsxChartKind.Column || chart.Kind == XlsxChartKind.Bar || chart.Kind == XlsxChartKind.Area;
        var value = Range(chart.YAxis, yMin[0] == double.MaxValue ? 0 : yMin[0], yMax[0] == double.MinValue ? 1 : yMax[0], columns);
        if (percent && chart.YAxis.NumberFormat == null) { value.Format = "0%"; }
        var second = hasSecondary ? Range(chart.SecondaryYAxis, yMin[1] == double.MaxValue ? 0 : yMin[1], yMax[1] == double.MinValue ? 1 : yMax[1], false) : null;
        var across = scatter ? Range(chart.XAxis, xMin == double.MaxValue ? 0 : xMin, xMax == double.MinValue ? 1 : xMax, false) : null;

        // Room for the labels and titles of each axis.
        var valueAxis = horizontal ? chart.XAxis : chart.YAxis;
        var valueLabels = new List<string>();
        foreach (var tick in value.Ticks()) { valueLabels.Add(Number(tick, value.Format, value.Step)); }
        var widest = 0.0;
        var verticalLabels = horizontal ? labels : valueLabels;
        foreach (var label in verticalLabels) { widest = Math.Max(widest, regular.Measure(label, text)); }
        var yTitle = horizontal ? chart.YAxis.Title : chart.YAxis.Title;
        var xTitle = chart.XAxis.Title;
        var plotLeft = left + widest + 6 + (string.IsNullOrEmpty(yTitle) ? 0 : look.AxisTitleSize * 1.6);
        var plotBottom = bottom - (text * 1.6) - (string.IsNullOrEmpty(xTitle) ? 0 : look.AxisTitleSize * 1.6);
        var plotRight = right - (text * 0.6);
        if (second != null)
        {
            var secondWidest = 0.0;
            foreach (var tick in second.Ticks()) { secondWidest = Math.Max(secondWidest, regular.Measure(Number(tick, second.Format, second.Step), text)); }
            plotRight = right - secondWidest - 6 - (string.IsNullOrEmpty(chart.SecondaryYAxis.Title) ? 0 : look.AxisTitleSize * 1.6);
        }

        var plotTop = top + (text * 0.5);
        double plotWidth = plotRight - plotLeft, plotHeight = plotBottom - plotTop;
        if (plotWidth < 10 || plotHeight < 10)
        {
            return;
        }

        double ValueX(Scale s, double v) => plotLeft + (s.Fraction(v) * plotWidth);
        double ValueY(Scale s, double v) => plotBottom - (s.Fraction(v) * plotHeight);

        // Gridlines and value labels.
        var grid = valueAxis.Gridlines ?? true;
        foreach (var tick in value.Ticks())
        {
            var label = Number(tick, value.Format, value.Step);
            if (horizontal)
            {
                var gx = ValueX(value, tick);
                if (grid) { page.DrawLine(gx, plotTop, gx, plotBottom, look.GridColor, 0.5); }
                page.DrawText(label, gx, plotBottom + text + 2, regular, text, look.TextColor, PdfAlign.Center);
            }
            else
            {
                var gy = ValueY(value, tick);
                if (grid) { page.DrawLine(plotLeft, gy, plotRight, gy, look.GridColor, 0.5); }
                page.DrawText(label, plotLeft - 5, gy + (text * 0.35), regular, text, look.TextColor, PdfAlign.Right);
            }
        }

        if (second != null)
        {
            foreach (var tick in second.Ticks())
            {
                var gy = ValueY(second, tick);
                page.DrawText(Number(tick, second.Format, second.Step), plotRight + 5, gy + (text * 0.35), regular, text, look.TextColor);
            }
        }

        if (across != null)
        {
            foreach (var tick in across.Ticks())
            {
                var gx = ValueX(across, tick);
                if (chart.XAxis.Gridlines == true) { page.DrawLine(gx, plotTop, gx, plotBottom, look.GridColor, 0.5); }
                page.DrawText(Number(tick, across.Format, across.Step), gx, plotBottom + text + 2, regular, text, look.TextColor, PdfAlign.Center);
            }
        }
        else if (categories > 0)
        {
            // Category labels, thinned out when they would overlap.
            var slot = (horizontal ? plotHeight : plotWidth) / categories;
            var every = 1;
            if (!horizontal)
            {
                var longest = 0.0;
                foreach (var label in labels) { longest = Math.Max(longest, regular.Measure(label, text)); }
                every = Math.Max(1, (int)Math.Ceiling((longest + 4) / slot));
            }

            for (var c = 0; c < categories; c += every)
            {
                var label = c < labels.Count ? labels[c] : string.Empty;
                if (horizontal)
                {
                    // Bars list their first category at the bottom, as Excel does.
                    page.DrawText(label, plotLeft - 5, plotBottom - ((c + 0.5) * slot) + (text * 0.35), regular, text, look.TextColor, PdfAlign.Right);
                }
                else
                {
                    page.DrawText(label, plotLeft + ((c + 0.5) * slot), plotBottom + text + 2, regular, text, look.TextColor, PdfAlign.Center);
                }
            }
        }

        // The series, kept inside the plot.
        page.PushClip(plotLeft, plotTop, plotWidth, plotHeight);
        var slotWidth = categories > 0 ? (horizontal ? plotHeight : plotWidth) / categories : 0;
        var clustered = new List<Data>();
        foreach (var s in series)
        {
            if (s.Kind == XlsxChartKind.Column || s.Kind == XlsxChartKind.Bar) { clustered.Add(s); }
        }

        var gap = chart.GapWidth / 100.0;
        var barCount = stacked ? 1 : Math.Max(1, clustered.Count);
        var barWidth = slotWidth / (barCount + gap);
        var baseline = Math.Max(value.Low, Math.Min(value.High, 0));
        var positive = new double[Math.Max(1, categories)];
        var negative = new double[Math.Max(1, categories)];
        var totals = new double[Math.Max(1, categories)];
        if (percent)
        {
            for (var c = 0; c < categories; c++)
            {
                foreach (var s in series) { if (c < s.Y.Count && s.Y[c] is double v) { totals[c] += Math.Abs(v); } }
            }
        }

        var areaBase = new double[Math.Max(1, categories)];
        foreach (var s in series)
        {
            var scale = s.Source.SecondaryAxis && second != null ? second : value;
            double Share(int c, double v) => percent && totals[c] > 0 ? v / totals[c] : v;
            if (s.Kind == XlsxChartKind.Column || s.Kind == XlsxChartKind.Bar)
            {
                var index = clustered.IndexOf(s);
                for (var c = 0; c < categories && c < s.Y.Count; c++)
                {
                    if (!(s.Y[c] is double raw)) { continue; }
                    var v = Share(c, raw);
                    double from = baseline, to = v;
                    if (stacked)
                    {
                        if (v >= 0) { from = positive[c]; positive[c] += v; to = positive[c]; }
                        else { from = negative[c]; negative[c] += v; to = negative[c]; }
                    }

                    var offset = (c * slotWidth) + (gap * barWidth / 2) + ((stacked ? 0 : index) * barWidth);
                    if (horizontal)
                    {
                        double x1 = ValueX(scale, Math.Min(from, to)), x2 = ValueX(scale, Math.Max(from, to));
                        page.FillRectangle(Math.Min(x1, x2), plotBottom - offset - barWidth, Math.Abs(x2 - x1), barWidth, s.Color);
                    }
                    else
                    {
                        double y1 = ValueY(scale, Math.Max(from, to)), y2 = ValueY(scale, Math.Min(from, to));
                        page.FillRectangle(plotLeft + offset, Math.Min(y1, y2), barWidth, Math.Abs(y2 - y1), s.Color);
                    }

                    if ((chart.DataLabels & XlsxDataLabels.Value) != 0)
                    {
                        var label = Number(raw, chart.YAxis.NumberFormat, value.Step);
                        if (horizontal) { page.DrawText(label, ValueX(scale, to) + 3, plotBottom - offset - (barWidth / 2) + (text * 0.35), regular, text, look.TextColor); }
                        else { page.DrawText(label, plotLeft + offset + (barWidth / 2), ValueY(scale, to) - 3, regular, text, look.TextColor, PdfAlign.Center); }
                    }
                }

                continue;
            }

            // Points along x, or one per category centre.
            var points = new List<(double X, double Y)>();
            var runs = new List<List<(double X, double Y)>> { points };
            for (var i = 0; i < s.Y.Count; i++)
            {
                double? px = scatter ? (s.X[i] is double xv ? ValueX(across!, xv) : (double?)null) : plotLeft + ((i + 0.5) * slotWidth);
                if (!(s.Y[i] is double raw) || px == null)
                {
                    if (points.Count > 0) { points = new List<(double, double)>(); runs.Add(points); }
                    continue;
                }

                var v = Share(Math.Min(i, Math.Max(0, categories - 1)), raw);
                if (s.Kind == XlsxChartKind.Area && stacked && i < areaBase.Length)
                {
                    v += areaBase[i];
                }

                points.Add((px.Value, ValueY(scale, v)));
            }

            if (s.Kind == XlsxChartKind.Area)
            {
                foreach (var run in runs)
                {
                    if (run.Count < 2) { continue; }
                    var shape = new List<(double, double)>(run);
                    for (var i = run.Count - 1; i >= 0; i--)
                    {
                        var index = (int)Math.Round(((run[i].X - plotLeft) / slotWidth) - 0.5);
                        var under = stacked && index >= 0 && index < areaBase.Length ? areaBase[index] : baseline;
                        shape.Add((run[i].X, ValueY(scale, under)));
                    }

                    page.FillPolygon(shape, s.Color);
                }

                if (stacked)
                {
                    for (var i = 0; i < s.Y.Count && i < areaBase.Length; i++)
                    {
                        if (s.Y[i] is double raw) { areaBase[i] += Share(i, raw); }
                    }
                }

                continue;
            }

            if (s.Source.ShowLine)
            {
                foreach (var run in runs)
                {
                    page.DrawPolyline(run, s.Color, Math.Max(0.5, s.Width), s.Dash);
                }
            }

            if (s.Source.MarkerSize > 0)
            {
                var marker = s.Source.MarkerColor != null ? "#" + Divit.Packaging.OfficeText.Rgb(s.Source.MarkerColor) : s.Color;
                foreach (var run in runs)
                {
                    foreach (var (px, py) in run)
                    {
                        page.FillEllipse(px, py, s.Source.MarkerSize / 2.0, s.Source.MarkerSize / 2.0, marker);
                    }
                }
            }
        }

        page.PopClip();

        // Axis lines and titles.
        page.DrawLine(plotLeft, plotBottom, plotRight, plotBottom, look.AxisColor, 0.75);
        page.DrawLine(plotLeft, plotTop, plotLeft, plotBottom, look.AxisColor, 0.75);
        if (second != null)
        {
            page.DrawLine(plotRight, plotTop, plotRight, plotBottom, look.AxisColor, 0.75);
            if (!string.IsNullOrEmpty(chart.SecondaryYAxis.Title))
            {
                var t = chart.SecondaryYAxis.Title!;
                page.DrawTurnedText(t, right - (look.AxisTitleSize * 0.3), ((plotTop + plotBottom) / 2) - (regular.Measure(t, look.AxisTitleSize) / 2), regular, look.AxisTitleSize, look.TextColor, -90);
            }
        }

        if (!string.IsNullOrEmpty(xTitle))
        {
            page.DrawText(xTitle!, (plotLeft + plotRight) / 2, bottom - (look.AxisTitleSize * 0.3), regular, look.AxisTitleSize, look.TextColor, PdfAlign.Center);
        }

        if (!string.IsNullOrEmpty(yTitle))
        {
            var t = yTitle!;
            page.DrawTurnedText(t, left + look.AxisTitleSize, ((plotTop + plotBottom) / 2) + (regular.Measure(t, look.AxisTitleSize) / 2), regular, look.AxisTitleSize, look.TextColor, 90);
        }

    }

    private static int Count(List<Data> series)
    {
        var count = 0;
        foreach (var s in series)
        {
            count = Math.Max(count, s.Y.Count);
        }

        return count;
    }

    private static void Pie(PdfPage page, XlsxChart chart, Data s, PdfFont regular, double left, double top, double right, double bottom)
    {
        var look = chart.Look;
        var total = 0.0;
        foreach (var v in s.Y) { if (v is double value && value > 0) { total += value; } }
        if (total <= 0)
        {
            return;
        }

        double cx = (left + right) / 2, cy = (top + bottom) / 2, radius = Math.Min(right - left, bottom - top) / 2 * 0.92;
        var angle = (chart.FirstSliceAngle - 90) * Math.PI / 180;
        for (var i = 0; i < s.Y.Count; i++)
        {
            if (!(s.Y[i] is double value) || value <= 0) { continue; }
            var sweep = value / total * 2 * Math.PI;
            var color = look.Palette.Count == 0 ? s.Color : look.Palette[i % look.Palette.Count];
            var shape = new List<(double, double)> { (cx, cy) };
            var steps = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 90)));
            for (var k = 0; k <= steps; k++)
            {
                var a = angle + (sweep * k / steps);
                shape.Add((cx + (radius * Math.Cos(a)), cy + (radius * Math.Sin(a))));
            }

            page.FillPolygon(shape, color);
            page.DrawPolyline(shape, "#FFFFFF", 0.75);
            if (chart.DataLabels != XlsxDataLabels.None)
            {
                var mid = angle + (sweep / 2);
                var at = chart.Kind == XlsxChartKind.Doughnut ? radius * (1 + (chart.HoleSize / 100.0)) / 2 : radius * 0.65;
                var parts = new List<string>();
                if ((chart.DataLabels & XlsxDataLabels.Category) != 0) { parts.Add(i < s.Labels.Count ? s.Labels[i] : string.Empty); }
                if ((chart.DataLabels & XlsxDataLabels.Value) != 0) { parts.Add(Number(value, null, 0)); }
                if ((chart.DataLabels & XlsxDataLabels.Percent) != 0) { parts.Add(Number(value / total, "0%", 0)); }
                page.DrawText(string.Join(" ", parts), cx + (at * Math.Cos(mid)), cy + (at * Math.Sin(mid)) + (look.TextSize * 0.35), regular, look.TextSize, "#FFFFFF", PdfAlign.Center);
            }

            angle += sweep;
        }

        if (chart.Kind == XlsxChartKind.Doughnut)
        {
            var hole = radius * Math.Max(10, Math.Min(90, chart.HoleSize)) / 100.0;
            page.FillEllipse(cx, cy, hole, hole, "#FFFFFF");
        }
    }

    /// <summary>
    /// The legend under the plot, entries centred in rows: a line for lines, a box for filled series or slices.
    /// </summary>
    private static double Legend(PdfPage page, XlsxChart chart, List<Data> series, bool pie, PdfFont regular, double left, double right, double bottom)
    {
        var look = chart.Look;
        var size = look.TextSize;
        var entries = new List<(string Text, string Color, bool Line, PdfDash Dash)>();
        var dots = new HashSet<string>(StringComparer.Ordinal);
        if (pie)
        {
            for (var i = 0; i < series[0].Labels.Count; i++)
            {
                entries.Add((series[0].Labels[i], look.Palette.Count == 0 ? series[0].Color : look.Palette[i % look.Palette.Count], false, PdfDash.Solid));
            }
        }
        else
        {
            foreach (var s in series)
            {
                if (s.Source.ShowInLegend)
                {
                    var line = s.Kind != XlsxChartKind.Column && s.Kind != XlsxChartKind.Bar && s.Kind != XlsxChartKind.Area;
                    entries.Add((s.Source.Name, s.Source.ShowLine || !line ? s.Color : "#" + Divit.Packaging.OfficeText.Rgb(s.Source.MarkerColor ?? s.Color), line && s.Source.ShowLine, s.Dash));
                    if (line && !s.Source.ShowLine) { dots.Add(s.Source.Name); }
                }
            }
        }

        if (entries.Count == 0)
        {
            return bottom;
        }

        // Rows of entries that fit the width, bottom row last.
        var rows = new List<List<(string Text, string Color, bool Line, PdfDash Dash, double Width)>> { new List<(string, string, bool, PdfDash, double)>() };
        var used = 0.0;
        foreach (var (text, color, line, dash) in entries)
        {
            var width = 22 + regular.Measure(text, size) + 12;
            if (used + width > right - left && rows[rows.Count - 1].Count > 0)
            {
                rows.Add(new List<(string, string, bool, PdfDash, double)>());
                used = 0;
            }

            rows[rows.Count - 1].Add((text, color, line, dash, width));
            used += width;
        }

        var rowHeight = size * 1.6;
        var y = bottom - (rows.Count * rowHeight) + (size * 1.1);
        foreach (var row in rows)
        {
            var total = 0.0;
            foreach (var entry in row) { total += entry.Width; }
            var x = ((left + right) / 2) - (total / 2);
            foreach (var (text, color, line, dash, width) in row)
            {
                if (line) { page.DrawLine(x, y - (size * 0.35), x + 16, y - (size * 0.35), color, 2, dash); }
                else if (dots.Contains(text)) { page.FillEllipse(x + 8, y - (size * 0.35), 4, 4, color); }
                else { page.FillRectangle(x + 4, y - (size * 0.75), 8, 8, color); }
                page.DrawText(text, x + 22, y, regular, size, look.TextColor);
                x += width;
            }

            y += rowHeight;
        }

        return bottom - (rows.Count * rowHeight) - 4;
    }
}
