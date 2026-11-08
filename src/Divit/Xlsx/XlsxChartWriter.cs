namespace Divit.Xlsx;

using System.Collections.Generic;
using System.Globalization;
using System.Xml;

/// <summary>
/// XML parts of charts: the drawing layer of a sheet and the part of each chart.
/// </summary>
/// <remarks>
/// DrawingML (ECMA-376 parts 20–21). Elements follow the order of the standard; Excel drops a chart with a wrong order
/// and "repairs" the file. Series values are read from the sheet and written as a cache too, so readers draw without recalculating.
/// </remarks>
internal static class XlsxChartWriter
{
    #region Fields

    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string SheetDrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// EMUs per pixel (96 dpi).
    /// </summary>
    private const long EmuPerPixel = 9525;

    #endregion

    #region Drawing layer

    /// <summary>
    /// The drawing layer of a sheet: each chart is a frame anchored to its top-left cell, sized in pixels.
    /// </summary>
    /// <param name="w">The writer.</param>
    /// <param name="charts">Charts of the sheet; chart i is <c>rId(i+1)</c> in the layer relationships.</param>
    /// <param name="images">Images of the sheet; image j is <c>rId(chart count + j + 1)</c>.</param>
    public static void WriteDrawing(XmlWriter w, IReadOnlyList<XlsxChart> charts, IReadOnlyList<XlsxSheet.XlsxImage> images)
    {
        w.WriteStartElement("xdr", "wsDr", SheetDrawingNs);
        w.WriteAttributeString("xmlns", "a", null, DrawingNs);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        for (var i = 0; i < charts.Count; i++)
        {
            var chart = charts[i];
            Anchor(w, chart.Row, chart.Column, chart.Width, chart.Height);

            w.WriteStartElement("graphicFrame", SheetDrawingNs);
            w.WriteAttributeString("macro", string.Empty);
            w.WriteStartElement("nvGraphicFramePr", SheetDrawingNs);
            w.WriteStartElement("cNvPr", SheetDrawingNs);
            w.WriteAttributeString("id", Int(i + 2));
            w.WriteAttributeString("name", "Chart " + Int(i + 1));
            w.WriteEndElement();
            w.WriteElementString("cNvGraphicFramePr", SheetDrawingNs, null);
            w.WriteEndElement();
            w.WriteStartElement("xfrm", SheetDrawingNs);
            w.WriteStartElement("off", DrawingNs); w.WriteAttributeString("x", "0"); w.WriteAttributeString("y", "0"); w.WriteEndElement();
            w.WriteStartElement("ext", DrawingNs); w.WriteAttributeString("cx", "0"); w.WriteAttributeString("cy", "0"); w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("graphic", DrawingNs);
            w.WriteStartElement("graphicData", DrawingNs);
            w.WriteAttributeString("uri", ChartNs);
            w.WriteStartElement("c", "chart", ChartNs);
            w.WriteAttributeString("id", RelNs, "rId" + Int(i + 1));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteElementString("clientData", SheetDrawingNs, null);
            w.WriteEndElement();
        }

        for (var j = 0; j < images.Count; j++)
        {
            var image = images[j];
            var cx = (image.Width * EmuPerPixel).ToString(CultureInfo.InvariantCulture);
            var cy = (image.Height * EmuPerPixel).ToString(CultureInfo.InvariantCulture);
            if (image.LastRow > 0)
            {
                CellsAnchor(w, image.Row, image.Column, image.LastRow, image.LastColumn);
            }
            else
            {
                Anchor(w, image.Row, image.Column, image.Width, image.Height);
            }

            w.WriteStartElement("pic", SheetDrawingNs);
            w.WriteStartElement("nvPicPr", SheetDrawingNs);
            w.WriteStartElement("cNvPr", SheetDrawingNs);
            w.WriteAttributeString("id", Int(charts.Count + j + 2));
            w.WriteAttributeString("name", "Picture " + Int(j + 1));
            if (image.Description != null) { w.WriteAttributeString("descr", Divit.Packaging.OfficeText.Clean(image.Description)); }
            w.WriteEndElement();
            w.WriteStartElement("cNvPicPr", SheetDrawingNs);
            w.WriteStartElement("picLocks", DrawingNs);
            w.WriteAttributeString("noChangeAspect", "1");
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("blipFill", SheetDrawingNs);
            w.WriteStartElement("blip", DrawingNs);
            w.WriteAttributeString("embed", RelNs, "rId" + Int(charts.Count + j + 1));
            w.WriteEndElement();
            w.WriteStartElement("stretch", DrawingNs);
            w.WriteElementString("fillRect", DrawingNs, null);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("spPr", SheetDrawingNs);
            w.WriteStartElement("xfrm", DrawingNs);
            w.WriteStartElement("off", DrawingNs); w.WriteAttributeString("x", "0"); w.WriteAttributeString("y", "0"); w.WriteEndElement();
            w.WriteStartElement("ext", DrawingNs); w.WriteAttributeString("cx", cx); w.WriteAttributeString("cy", cy); w.WriteEndElement();
            w.WriteEndElement();
            w.WriteStartElement("prstGeom", DrawingNs);
            w.WriteAttributeString("prst", "rect");
            w.WriteElementString("avLst", DrawingNs, null);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteElementString("clientData", SheetDrawingNs, null);
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Position of a frame that fills a range: from the top-left corner of the first cell to the top-left corner of
    /// the cell after the last one (counted from 0), moving and sizing with the cells.
    /// </summary>
    private static void CellsAnchor(XmlWriter w, int row1, int column1, int row2, int column2)
    {
        w.WriteStartElement("twoCellAnchor", SheetDrawingNs);
        w.WriteAttributeString("editAs", "twoCell");
        Marker(w, "from", row1 - 1, column1 - 1);
        Marker(w, "to", row2, column2);
    }

    private static void Marker(XmlWriter w, string name, int row, int column)
    {
        w.WriteStartElement(name, SheetDrawingNs);
        w.WriteElementString("col", SheetDrawingNs, Int(column));
        w.WriteElementString("colOff", SheetDrawingNs, "0");
        w.WriteElementString("row", SheetDrawingNs, Int(row));
        w.WriteElementString("rowOff", SheetDrawingNs, "0");
        w.WriteEndElement();
    }

    /// <summary>
    /// Position of a frame: the top-left cell (counted from 0) and the size in EMUs.
    /// </summary>
    private static void Anchor(XmlWriter w, int row, int column, int width, int height)
    {
        w.WriteStartElement("oneCellAnchor", SheetDrawingNs);
        w.WriteStartElement("from", SheetDrawingNs);
        w.WriteElementString("col", SheetDrawingNs, Int(column - 1));
        w.WriteElementString("colOff", SheetDrawingNs, "0");
        w.WriteElementString("row", SheetDrawingNs, Int(row - 1));
        w.WriteElementString("rowOff", SheetDrawingNs, "0");
        w.WriteEndElement();
        w.WriteStartElement("ext", SheetDrawingNs);
        w.WriteAttributeString("cx", (width * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
        w.WriteAttributeString("cy", (height * EmuPerPixel).ToString(CultureInfo.InvariantCulture));
        w.WriteEndElement();
    }

    #endregion

    #region Chart

    /// <summary>
    /// The part of one chart: title, series, axes, legend and look.
    /// </summary>
    /// <param name="w">The writer.</param>
    /// <param name="chart">The chart.</param>
    /// <param name="defaultFont">The workbook font, used when the look names none.</param>
    /// <param name="workbook">Relationship id of the embedded workbook of a chart in Word; none in Excel.</param>
    public static void WriteChart(XmlWriter w, XlsxChart chart, string defaultFont, string? workbook = null)
    {
        var look = chart.Look;
        var font = look.FontName ?? defaultFont;

        w.WriteStartElement("c", "chartSpace", ChartNs);
        w.WriteAttributeString("xmlns", "a", null, DrawingNs);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        Val(w, "roundedCorners", "0");

        w.WriteStartElement("chart", ChartNs);
        if (chart.Title != null)
        {
            w.WriteStartElement("title", ChartNs);
            w.WriteStartElement("tx", ChartNs);
            w.WriteStartElement("rich", ChartNs);
            w.WriteElementString("bodyPr", DrawingNs, null);
            Paragraph(w, chart.Title, font, look.TitleSize, true, look.TitleColor);
            if (chart.Subtitle != null)
            {
                Paragraph(w, chart.Subtitle, font, look.SubtitleSize, false, look.SubtitleColor);
            }

            w.WriteEndElement();
            w.WriteEndElement();
            Val(w, "overlay", "0");
            w.WriteEndElement();
            Val(w, "autoTitleDeleted", "0");
        }
        else
        {
            Val(w, "autoTitleDeleted", "1");
        }

        var polar = chart.Kind == XlsxChartKind.Polar;
        w.WriteStartElement("plotArea", ChartNs);
        if (chart.PlotLayout is (double, double, double, double) layout)
        {
            // Manual layout: "inner" sets the area inside the axes; labels and titles stay outside it.
            w.WriteStartElement("layout", ChartNs);
            w.WriteStartElement("manualLayout", ChartNs);
            Val(w, "layoutTarget", "inner");
            Val(w, "xMode", "edge");
            Val(w, "yMode", "edge");
            Val(w, "x", layout.X.ToString("0.####", CultureInfo.InvariantCulture));
            Val(w, "y", layout.Y.ToString("0.####", CultureInfo.InvariantCulture));
            Val(w, "w", layout.Width.ToString("0.####", CultureInfo.InvariantCulture));
            Val(w, "h", layout.Height.ToString("0.####", CultureInfo.InvariantCulture));
            w.WriteEndElement();
            w.WriteEndElement();
        }
        else
        {
            w.WriteElementString("layout", ChartNs, null);
        }

        var pie = chart.Kind == XlsxChartKind.Pie || chart.Kind == XlsxChartKind.Doughnut;
        if (chart.Kind == XlsxChartKind.Scatter || polar)
        {
            // Series on the right axis form a second scatter group with its own pair of axes.
            var secondary = !polar && HasSecondary(chart);
            foreach (var right in secondary ? new[] { false, true } : new[] { false })
            {
                w.WriteStartElement(polar ? "radarChart" : "scatterChart", ChartNs);
                Val(w, polar ? "radarStyle" : "scatterStyle", polar ? "marker" : "smoothMarker");
                Val(w, "varyColors", "0");
                for (var i = 0; i < chart.Series.Count; i++)
                {
                    if (polar || chart.Series[i].SecondaryAxis == right)
                    {
                        Series(w, chart, chart.Series[i], i, polar);
                    }
                }

                DataLabels(w, chart.DataLabels, false, font, look);
                Val(w, "axId", right ? "3" : "1");
                Val(w, "axId", right ? "4" : "2");
                w.WriteEndElement();
            }

            if (polar)
            {
                AngleAxis(w, chart, font, look);
                Axis(w, chart.YAxis, 2, 1, "l", font, look, AxisMode.Radius);
            }
            else
            {
                Axis(w, chart.XAxis, 1, 2, "b", font, look, AxisMode.Scatter);
                Axis(w, chart.YAxis, 2, 1, "l", font, look, AxisMode.Scatter);
                if (secondary)
                {
                    // The second X axis follows the first and stays hidden; the second Y axis stands on the right.
                    Axis(w, chart.XAxis, 3, 4, "b", font, look, AxisMode.Scatter, deleted: true);
                    Axis(w, chart.SecondaryYAxis, 4, 3, "r", font, look, AxisMode.Scatter, crosses: "max");
                }
            }
        }
        else if (pie)
        {
            PiePlot(w, chart, font, look);
        }
        else
        {
            CategoryPlot(w, chart, font, look);
        }

        w.WriteEndElement();

        var legend = chart.ShowLegend ?? (pie || chart.Series.Count > 1);
        if (legend)
        {
            w.WriteStartElement("legend", ChartNs);
            Val(w, "legendPos", "b");

            // Series hidden from the legend (right after legendPos, as the standard orders it).
            for (var i = 0; i < chart.Series.Count; i++)
            {
                if (!chart.Series[i].ShowInLegend)
                {
                    w.WriteStartElement("legendEntry", ChartNs);
                    Val(w, "idx", Int(i));
                    Val(w, "delete", "1");
                    w.WriteEndElement();
                }
            }
            Val(w, "overlay", "0");
            TextProperties(w, font, look.TextSize, false, look.TextColor);
            w.WriteEndElement();
        }

        // Hidden columns are drawn too: helper series live in hidden columns.
        Val(w, "plotVisOnly", "0");
        Val(w, "dispBlanksAs", "gap");
        w.WriteEndElement();

        // Chart area: white background with a thin border (or none).
        w.WriteStartElement("spPr", ChartNs);
        Solid(w, "#FFFFFF");
        if (look.BorderColor != null)
        {
            Line(w, look.BorderColor, 0.75, "solid");
        }
        else
        {
            w.WriteStartElement("ln", DrawingNs); w.WriteElementString("noFill", DrawingNs, null); w.WriteEndElement();
        }

        w.WriteEndElement();
        TextProperties(w, font, look.TextSize, false, look.TextColor);
        if (workbook != null)
        {
            // The workbook embedded with a chart in Word, where "Edit Data" opens it.
            w.WriteStartElement("externalData", ChartNs);
            w.WriteAttributeString("r", "id", RelNs, workbook);
            Val(w, "autoUpdate", "0");
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private static void Series(XmlWriter w, XlsxChart chart, XlsxSeries s, int index, bool polar)
    {
        w.WriteStartElement("ser", ChartNs);
        Val(w, "idx", Int(index));
        Val(w, "order", Int(index));
        w.WriteStartElement("tx", ChartNs);
        w.WriteElementString("v", ChartNs, s.Name);
        w.WriteEndElement();
        w.WriteStartElement("spPr", ChartNs);
        if (s.ShowLine)
        {
            Line(w, s.Line.Color, s.Line.Width, s.Line.Dash);
        }
        else
        {
            w.WriteStartElement("ln", DrawingNs); w.WriteElementString("noFill", DrawingNs, null); w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteStartElement("marker", ChartNs);
        if (s.MarkerSize > 0)
        {
            var color = s.MarkerColor ?? s.Line.Color;
            Val(w, "symbol", "circle");
            Val(w, "size", Int(s.MarkerSize));
            w.WriteStartElement("spPr", ChartNs);
            Solid(w, color);
            Line(w, color, 0.75, "solid");
            w.WriteEndElement();
        }
        else
        {
            Val(w, "symbol", "none");
        }

        w.WriteEndElement();
        if (!polar)
        {
            TrendAndErrors(w, chart.Sheet, s, s.Line.Color, true);
        }

        // Polar (radar): angles are categories and values the radius; scatter: X and Y.
        Values(w, polar ? "cat" : "xVal", chart.Sheet, s.XColumn, s.FirstRow, s.LastRow);
        Values(w, polar ? "val" : "yVal", chart.Sheet, s.YColumn, s.FirstRow, s.LastRow);
        if (!polar)
        {
            Val(w, "smooth", s.Smooth && s.ShowLine ? "1" : "0");
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Angle axis of a polar chart (the category axis of Excel's radar chart): labels are the angles of the X column;
    /// <see cref="XlsxAxis.MajorUnit"/> is the label step in degrees (for example 30) and becomes a skip count.
    /// </summary>
    private static void AngleAxis(XmlWriter w, XlsxChart chart, string font, XlsxChartLook look)
    {
        var skip = 1;
        if (chart.Series.Count > 0 && chart.XAxis.MajorUnit != null)
        {
            var s = chart.Series[0];
            var a = chart.Sheet.NumberAt(s.FirstRow, s.XColumn);
            var b = chart.Sheet.NumberAt(s.FirstRow + 1, s.XColumn);
            if (a != null && b != null && b.Value != a.Value)
            {
                skip = System.Math.Max(1, (int)System.Math.Round(chart.XAxis.MajorUnit.Value / System.Math.Abs(b.Value - a.Value)));
            }
        }

        w.WriteStartElement("catAx", ChartNs);
        Val(w, "axId", "1");
        w.WriteStartElement("scaling", ChartNs);
        Val(w, "orientation", "minMax");
        w.WriteEndElement();
        Val(w, "delete", "0");
        Val(w, "axPos", "b");
        w.WriteStartElement("majorGridlines", ChartNs);
        w.WriteStartElement("spPr", ChartNs);
        Line(w, look.GridColor, 0.75, "solid");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("numFmt", ChartNs);
        w.WriteAttributeString("formatCode", chart.XAxis.NumberFormat ?? "0\"°\"");
        w.WriteAttributeString("sourceLinked", "0");
        w.WriteEndElement();
        Val(w, "majorTickMark", "none");
        Val(w, "minorTickMark", "none");
        Val(w, "tickLblPos", "nextTo");
        w.WriteStartElement("spPr", ChartNs);
        Line(w, look.AxisColor, 0.75, "solid");
        w.WriteEndElement();
        TextProperties(w, font, look.TextSize, false, look.TextColor);
        Val(w, "crossAx", "2");
        Val(w, "crosses", "autoZero");
        Val(w, "auto", "1");
        Val(w, "lblAlgn", "ctr");
        Val(w, "lblOffset", "100");
        Val(w, "tickLblSkip", Int(skip));
        Val(w, "tickMarkSkip", Int(skip));
        Val(w, "noMultiLvlLbl", "0");
        w.WriteEndElement();
    }

    /// <summary>
    /// A column range: the reference (<c>'Sheet'!$A$7:$A$79</c>) plus a cache of its values.
    /// </summary>
    private static void Values(XmlWriter w, string name, XlsxSheet sheet, int column, int firstRow, int lastRow, string? format = null)
    {
        var col = "$" + XlsxWorkbook.ColumnName(column) + "$";
        w.WriteStartElement(name, ChartNs);
        w.WriteStartElement("numRef", ChartNs);
        w.WriteElementString("f", ChartNs, "'" + sheet.Name.Replace("'", "''") + "'!" + col + Int(firstRow) + ":" + col + Int(lastRow));
        w.WriteStartElement("numCache", ChartNs);
        w.WriteElementString("formatCode", ChartNs, format ?? "General");
        Val(w, "ptCount", Int(lastRow - firstRow + 1));
        for (var r = firstRow; r <= lastRow; r++)
        {
            if (sheet.NumberAt(r, column) is double v)
            {
                w.WriteStartElement("pt", ChartNs);
                w.WriteAttributeString("idx", Int(r - firstRow));
                w.WriteElementString("v", ChartNs, v.ToString("R", CultureInfo.InvariantCulture));
                w.WriteEndElement();
            }
        }

        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The role of a value axis, which decides where it crosses the other axis and where its labels go.
    /// </summary>
    private enum AxisMode
    {
        /// <summary>
        /// An axis of a scatter chart: crosses the other at its minimum, so the axes stay in the bottom-left corner
        /// even when every value is negative.
        /// </summary>
        Scatter,

        /// <summary>
        /// The radius of a polar chart.
        /// </summary>
        Radius,

        /// <summary>
        /// The value axis of a column, bar, line or area chart: crosses at zero, columns sit between the ticks.
        /// </summary>
        Value,
    }

    /// <summary>
    /// A value axis.
    /// </summary>
    private static void Axis(XmlWriter w, XlsxAxis axis, int id, int crossId, string position, string font, XlsxChartLook look, AxisMode mode,
        string? defaultFormat = null, bool deleted = false, string? crosses = null)
    {
        w.WriteStartElement("valAx", ChartNs);
        Val(w, "axId", Int(id));
        Scaling(w, axis, true);
        Val(w, "delete", deleted ? "1" : "0");
        Val(w, "axPos", position);
        // A hidden axis only carries the scale for its series: no gridlines, no title.
        if (!deleted && (axis.Gridlines ?? true))
        {
            Gridlines(w, look);
        }

        if (!deleted) { AxisTitle(w, axis, position, font, look); }
        var format = axis.NumberFormat ?? defaultFormat;
        w.WriteStartElement("numFmt", ChartNs);
        w.WriteAttributeString("formatCode", format ?? "General");
        w.WriteAttributeString("sourceLinked", format == null ? "1" : "0");
        w.WriteEndElement();
        Val(w, "majorTickMark", mode == AxisMode.Radius ? "cross" : "out");
        Val(w, "minorTickMark", "none");
        // An axis crossing at the far end (the right axis) keeps its labels beside it; "low" would move them to the left.
        Val(w, "tickLblPos", mode == AxisMode.Scatter && crosses == null ? "low" : "nextTo");
        w.WriteStartElement("spPr", ChartNs);
        Line(w, look.AxisColor, 0.75, "solid");
        w.WriteEndElement();
        TextProperties(w, font, look.TextSize, false, look.TextColor);
        Val(w, "crossAx", Int(crossId));
        Val(w, "crosses", crosses ?? (mode == AxisMode.Scatter ? "min" : "autoZero"));
        Val(w, "crossBetween", mode == AxisMode.Scatter ? "midCat" : "between");
        if (axis.MajorUnit != null) { Val(w, "majorUnit", Num(axis.MajorUnit.Value)); }
        w.WriteEndElement();
    }

    /// <summary>
    /// The category axis of a column, bar, line or area chart. Labels sit at the low end, so they stay at the edge
    /// when values are negative.
    /// </summary>
    private static void CategoryAxis(XmlWriter w, XlsxAxis axis, int id, int crossId, string position, string font, XlsxChartLook look, bool deleted = false)
    {
        w.WriteStartElement("catAx", ChartNs);
        Val(w, "axId", Int(id));
        Scaling(w, axis, false);
        Val(w, "delete", deleted ? "1" : "0");
        Val(w, "axPos", position);
        if (!deleted && (axis.Gridlines ?? false))
        {
            Gridlines(w, look);
        }

        if (!deleted) { AxisTitle(w, axis, position, font, look); }
        w.WriteStartElement("numFmt", ChartNs);
        w.WriteAttributeString("formatCode", axis.NumberFormat ?? "General");
        w.WriteAttributeString("sourceLinked", axis.NumberFormat == null ? "1" : "0");
        w.WriteEndElement();
        Val(w, "majorTickMark", "out");
        Val(w, "minorTickMark", "none");
        Val(w, "tickLblPos", "low");
        w.WriteStartElement("spPr", ChartNs);
        Line(w, look.AxisColor, 0.75, "solid");
        w.WriteEndElement();
        TextProperties(w, font, look.TextSize, false, look.TextColor);
        Val(w, "crossAx", Int(crossId));
        Val(w, "crosses", "autoZero");
        Val(w, "auto", "1");
        Val(w, "lblAlgn", "ctr");
        Val(w, "lblOffset", "100");
        Val(w, "noMultiLvlLbl", "0");
        w.WriteEndElement();
    }

    /// <summary>
    /// Scale of an axis: logarithmic base, direction, and on value axes the range; in the order of the standard.
    /// </summary>
    private static void Scaling(XmlWriter w, XlsxAxis axis, bool values)
    {
        w.WriteStartElement("scaling", ChartNs);
        if (values && axis.LogBase is double logBase)
        {
            Val(w, "logBase", Num(System.Math.Max(2, System.Math.Min(1000, logBase))));
        }

        Val(w, "orientation", axis.Reverse ? "maxMin" : "minMax");
        if (values && axis.Max != null) { Val(w, "max", Num(axis.Max.Value)); }
        if (values && axis.Min != null) { Val(w, "min", Num(axis.Min.Value)); }
        w.WriteEndElement();
    }

    private static bool HasSecondary(XlsxChart chart)
    {
        foreach (var s in chart.Series)
        {
            if (s.SecondaryAxis)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The trendline and error bars of a series, between its labels and its values.
    /// </summary>
    private static void TrendAndErrors(XmlWriter w, XlsxSheet sheet, XlsxSeries s, string color, bool scatter)
    {
        if (s.Trendline is XlsxTrendline trend)
        {
            w.WriteStartElement("trendline", ChartNs);
            if (trend.Name != null) { w.WriteElementString("name", ChartNs, Divit.Packaging.OfficeText.Clean(trend.Name)); }
            var line = trend.Line ?? new XlsxLine(color, 1.25, "sysDot");
            w.WriteStartElement("spPr", ChartNs);
            Line(w, line.Color, line.Width, line.Dash);
            w.WriteEndElement();
            Val(w, "trendlineType", TrendName(trend.Kind));
            if (trend.Kind == XlsxTrendlineKind.Polynomial) { Val(w, "order", Int(Clamp(trend.Order, 2, 6))); }
            if (trend.Kind == XlsxTrendlineKind.MovingAverage) { Val(w, "period", Int(System.Math.Max(2, trend.Period))); }
            Val(w, "dispRSqr", trend.ShowRSquared ? "1" : "0");
            Val(w, "dispEq", trend.ShowEquation ? "1" : "0");
            w.WriteEndElement();
        }

        if (s.ErrorBars is XlsxErrorBars bars)
        {
            w.WriteStartElement("errBars", ChartNs);

            // Scatter charts have X and Y error bars and must say which; the other charts have Y only.
            if (scatter) { Val(w, "errDir", "y"); }
            Val(w, "errBarType", bars.Direction == XlsxErrorDirection.Plus ? "plus" : bars.Direction == XlsxErrorDirection.Minus ? "minus" : "both");
            Val(w, "errValType", bars.Type);
            Val(w, "noEndCap", bars.EndCap ? "0" : "1");
            if (bars.Type == "cust")
            {
                Values(w, "plus", sheet, bars.PlusColumn, s.FirstRow, s.LastRow);
                Values(w, "minus", sheet, bars.MinusColumn, s.FirstRow, s.LastRow);
            }
            else if (bars.Type != "stdErr")
            {
                Val(w, "val", Num(bars.Value));
            }

            var line = bars.Line ?? new XlsxLine("#404040", 1, "solid");
            w.WriteStartElement("spPr", ChartNs);
            Line(w, line.Color, line.Width, line.Dash);
            w.WriteEndElement();
            w.WriteEndElement();
        }
    }

    private static string TrendName(XlsxTrendlineKind kind)
    {
        switch (kind)
        {
            case XlsxTrendlineKind.Exponential: return "exp";
            case XlsxTrendlineKind.Logarithmic: return "log";
            case XlsxTrendlineKind.Polynomial: return "poly";
            case XlsxTrendlineKind.Power: return "power";
            case XlsxTrendlineKind.MovingAverage: return "movingAvg";
            default: return "linear";
        }
    }

    private static void Gridlines(XmlWriter w, XlsxChartLook look)
    {
        w.WriteStartElement("majorGridlines", ChartNs);
        w.WriteStartElement("spPr", ChartNs);
        Line(w, look.GridColor, 0.75, "solid");
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// An axis title; turned upright on a vertical axis.
    /// </summary>
    private static void AxisTitle(XmlWriter w, XlsxAxis axis, string position, string font, XlsxChartLook look)
    {
        if (axis.Title == null)
        {
            return;
        }

        w.WriteStartElement("title", ChartNs);
        w.WriteStartElement("tx", ChartNs);
        w.WriteStartElement("rich", ChartNs);
        w.WriteStartElement("bodyPr", DrawingNs);
        if (position == "l") { w.WriteAttributeString("rot", "-5400000"); w.WriteAttributeString("vert", "horz"); }
        w.WriteEndElement();
        Paragraph(w, axis.Title, font, look.AxisTitleSize, true, look.TextColor);
        w.WriteEndElement();
        w.WriteEndElement();
        Val(w, "overlay", "0");
        w.WriteEndElement();
    }

    /// <summary>
    /// Columns, bars, lines or areas over categories, with a category axis and a value axis. Series of another kind
    /// or on the right axis form groups of their own, one chart element each, in the order they first appear; the
    /// right axis gets a hidden category axis of its own.
    /// </summary>
    private static void CategoryPlot(XmlWriter w, XlsxChart chart, string font, XlsxChartLook look)
    {
        var groups = new List<(XlsxChartKind Kind, bool Right)>();
        foreach (var s in chart.Series)
        {
            var group = (KindOf(chart, s), s.SecondaryAxis);
            if (!groups.Contains(group))
            {
                groups.Add(group);
            }
        }

        if (groups.Count == 0)
        {
            groups.Add((chart.Kind, false));
        }

        foreach (var group in groups)
        {
            CategoryGroup(w, chart, group.Kind, group.Right, font, look);
        }

        // A bar chart turns the axes: categories on the left, values below.
        var horizontal = chart.Kind == XlsxChartKind.Bar;
        var percent = chart.Grouping == XlsxChartGrouping.PercentStacked ? "0%" : null;
        CategoryAxis(w, chart.XAxis, 1, 2, horizontal ? "l" : "b", font, look);

        // Reversed categories move the value axis to the far side; crossing at the last category keeps it at the edge
        // it started at.
        Axis(w, chart.YAxis, 2, 1, horizontal ? "b" : "l", font, look, AxisMode.Value, percent, crosses: chart.XAxis.Reverse ? "max" : null);
        if (groups.Exists(g => g.Right))
        {
            CategoryAxis(w, chart.XAxis, 3, 4, horizontal ? "l" : "b", font, look, deleted: true);
            Axis(w, chart.SecondaryYAxis, 4, 3, horizontal ? "t" : "r", font, look, AxisMode.Value, percent, crosses: "max");
        }
    }

    /// <summary>
    /// The kind a series is drawn in: its own column, line or area kind in a column, line or area chart, else the
    /// chart's.
    /// </summary>
    private static XlsxChartKind KindOf(XlsxChart chart, XlsxSeries s)
    {
        var combinable = chart.Kind == XlsxChartKind.Column || chart.Kind == XlsxChartKind.Line || chart.Kind == XlsxChartKind.Area;
        var own = s.Kind;
        return combinable && (own == XlsxChartKind.Column || own == XlsxChartKind.Line || own == XlsxChartKind.Area) ? own!.Value : chart.Kind;
    }

    private static void CategoryGroup(XmlWriter w, XlsxChart chart, XlsxChartKind kind, bool right, string font, XlsxChartLook look)
    {
        var bars = kind == XlsxChartKind.Column || kind == XlsxChartKind.Bar;
        var element = bars ? "barChart" : kind == XlsxChartKind.Line ? "lineChart" : "areaChart";
        w.WriteStartElement(element, ChartNs);
        if (bars)
        {
            Val(w, "barDir", kind == XlsxChartKind.Bar ? "bar" : "col");
        }

        // Bars call side by side "clustered", lines and areas "standard".
        var grouping = chart.Grouping == XlsxChartGrouping.Stacked ? "stacked"
            : chart.Grouping == XlsxChartGrouping.PercentStacked ? "percentStacked"
            : bars ? "clustered" : "standard";
        Val(w, "grouping", grouping);
        Val(w, "varyColors", "0");
        for (var i = 0; i < chart.Series.Count; i++)
        {
            var s = chart.Series[i];
            if (KindOf(chart, s) == kind && s.SecondaryAxis == right)
            {
                CategorySeries(w, chart, s, i, kind);
            }
        }

        DataLabels(w, chart.DataLabels, false, font, look);
        if (bars)
        {
            Val(w, "gapWidth", Int(Clamp(chart.GapWidth, 0, 500)));

            // Stacked columns lie on each other completely.
            if (chart.Grouping != XlsxChartGrouping.Standard) { Val(w, "overlap", "100"); }
        }
        else if (kind == XlsxChartKind.Line)
        {
            Val(w, "marker", "1");
        }

        Val(w, "axId", right ? "3" : "1");
        Val(w, "axId", right ? "4" : "2");
        w.WriteEndElement();
    }

    private static void CategorySeries(XmlWriter w, XlsxChart chart, XlsxSeries s, int index, XlsxChartKind kind)
    {
        var look = chart.Look;
        var color = s.OwnColor ? s.Line.Color : PaletteColor(look, index);
        w.WriteStartElement("ser", ChartNs);
        Val(w, "idx", Int(index));
        Val(w, "order", Int(index));
        w.WriteStartElement("tx", ChartNs);
        w.WriteElementString("v", ChartNs, s.Name);
        w.WriteEndElement();
        w.WriteStartElement("spPr", ChartNs);
        if (kind == XlsxChartKind.Line)
        {
            Line(w, color, s.OwnColor ? s.Line.Width : look.LineWidth, s.OwnColor ? s.Line.Dash : "solid");
        }
        else
        {
            Solid(w, color);
            w.WriteStartElement("ln", DrawingNs); w.WriteElementString("noFill", DrawingNs, null); w.WriteEndElement();
        }

        w.WriteEndElement();
        if (kind == XlsxChartKind.Column || kind == XlsxChartKind.Bar)
        {
            Val(w, "invertIfNegative", "0");
        }

        if (kind == XlsxChartKind.Line)
        {
            Marker(w, s, color);
        }

        TrendAndErrors(w, chart.Sheet, s, color, false);
        Categories(w, chart.Sheet, s.XColumn, s.FirstRow, s.LastRow);
        Values(w, "val", chart.Sheet, s.YColumn, s.FirstRow, s.LastRow);
        if (kind == XlsxChartKind.Line)
        {
            Val(w, "smooth", s.Smooth ? "1" : "0");
        }

        w.WriteEndElement();
    }

    private static void Marker(XmlWriter w, XlsxSeries s, string lineColor)
    {
        w.WriteStartElement("marker", ChartNs);
        if (s.MarkerSize > 0)
        {
            var color = s.MarkerColor ?? lineColor;
            Val(w, "symbol", "circle");
            Val(w, "size", Int(Clamp(s.MarkerSize, 2, 72)));
            w.WriteStartElement("spPr", ChartNs);
            Solid(w, color);
            Line(w, color, 0.75, "solid");
            w.WriteEndElement();
        }
        else
        {
            Val(w, "symbol", "none");
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Pie (first series) or doughnut (a ring per series): each slice takes the next palette colour, with a thin
    /// white line between slices.
    /// </summary>
    private static void PiePlot(XmlWriter w, XlsxChart chart, string font, XlsxChartLook look)
    {
        var doughnut = chart.Kind == XlsxChartKind.Doughnut;
        w.WriteStartElement(doughnut ? "doughnutChart" : "pieChart", ChartNs);
        Val(w, "varyColors", "1");
        for (var i = 0; i < chart.Series.Count; i++)
        {
            var s = chart.Series[i];
            w.WriteStartElement("ser", ChartNs);
            Val(w, "idx", Int(i));
            Val(w, "order", Int(i));
            w.WriteStartElement("tx", ChartNs);
            w.WriteElementString("v", ChartNs, s.Name);
            w.WriteEndElement();
            if (s.Explosion > 0) { Val(w, "explosion", Int(Clamp(s.Explosion, 0, 400))); }
            for (var p = 0; p <= s.LastRow - s.FirstRow; p++)
            {
                w.WriteStartElement("dPt", ChartNs);
                Val(w, "idx", Int(p));
                Val(w, "bubble3D", "0");
                w.WriteStartElement("spPr", ChartNs);
                Solid(w, PaletteColor(look, p));
                Line(w, "#FFFFFF", 0.75, "solid");
                w.WriteEndElement();
                w.WriteEndElement();
            }

            Categories(w, chart.Sheet, s.XColumn, s.FirstRow, s.LastRow);
            Values(w, "val", chart.Sheet, s.YColumn, s.FirstRow, s.LastRow);
            w.WriteEndElement();
        }

        DataLabels(w, chart.DataLabels, true, font, look);
        Val(w, "firstSliceAng", Int(Clamp(chart.FirstSliceAngle, 0, 360)));
        if (doughnut)
        {
            Val(w, "holeSize", Int(Clamp(chart.HoleSize, 10, 90)));
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Labels on the points, for all series of the chart; nothing when none are asked for.
    /// </summary>
    private static void DataLabels(XmlWriter w, XlsxDataLabels labels, bool pie, string font, XlsxChartLook look)
    {
        if (labels == XlsxDataLabels.None)
        {
            return;
        }

        w.WriteStartElement("dLbls", ChartNs);
        TextProperties(w, font, look.TextSize, false, look.TextColor);
        Val(w, "showLegendKey", "0");
        Val(w, "showVal", (labels & XlsxDataLabels.Value) != 0 ? "1" : "0");
        Val(w, "showCatName", (labels & XlsxDataLabels.Category) != 0 ? "1" : "0");
        Val(w, "showSerName", (labels & XlsxDataLabels.SeriesName) != 0 ? "1" : "0");
        Val(w, "showPercent", pie && (labels & XlsxDataLabels.Percent) != 0 ? "1" : "0");
        Val(w, "showBubbleSize", "0");
        if (pie) { Val(w, "showLeaderLines", "1"); }
        w.WriteEndElement();
    }

    /// <summary>
    /// The categories of a series: text when any cell holds text, otherwise numbers (or dates) in the format of the
    /// first cell.
    /// </summary>
    private static void Categories(XmlWriter w, XlsxSheet sheet, int column, int firstRow, int lastRow)
    {
        var text = false;
        for (var r = firstRow; r <= lastRow && !text; r++)
        {
            text = sheet.NumberAt(r, column) == null && sheet.TextAt(r, column) != null;
        }

        if (!text)
        {
            Values(w, "cat", sheet, column, firstRow, lastRow, sheet.FormatAt(firstRow, column));
            return;
        }

        var col = "$" + XlsxWorkbook.ColumnName(column) + "$";
        w.WriteStartElement("cat", ChartNs);
        w.WriteStartElement("strRef", ChartNs);
        w.WriteElementString("f", ChartNs, "'" + sheet.Name.Replace("'", "''") + "'!" + col + Int(firstRow) + ":" + col + Int(lastRow));
        w.WriteStartElement("strCache", ChartNs);
        Val(w, "ptCount", Int(lastRow - firstRow + 1));
        for (var r = firstRow; r <= lastRow; r++)
        {
            if (sheet.TextAt(r, column) is string value)
            {
                w.WriteStartElement("pt", ChartNs);
                w.WriteAttributeString("idx", Int(r - firstRow));
                w.WriteElementString("v", ChartNs, Divit.Packaging.OfficeText.Clean(value));
                w.WriteEndElement();
            }
        }

        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static string PaletteColor(XlsxChartLook look, int index) =>
        look.Palette.Count == 0 ? look.SeriesLine.Color : look.Palette[index % look.Palette.Count];

    private static int Clamp(int value, int min, int max) => System.Math.Max(min, System.Math.Min(max, value));

    #endregion

    #region Helpers

    private static void Paragraph(XmlWriter w, string text, string font, double size, bool bold, string color)
    {
        w.WriteStartElement("p", DrawingNs);
        w.WriteStartElement("pPr", DrawingNs);
        RunProperties(w, "defRPr", font, size, bold, color);
        w.WriteEndElement();
        w.WriteStartElement("r", DrawingNs);
        RunProperties(w, "rPr", font, size, bold, color);
        w.WriteElementString("t", DrawingNs, text);
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void RunProperties(XmlWriter w, string element, string font, double size, bool bold, string color)
    {
        w.WriteStartElement(element, DrawingNs);
        w.WriteAttributeString("sz", Int((int)(size * 100)));
        w.WriteAttributeString("b", bold ? "1" : "0");
        Solid(w, color);
        w.WriteStartElement("latin", DrawingNs);
        w.WriteAttributeString("typeface", font);
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// Default text properties of an element (axis labels, legend, the whole chart).
    /// </summary>
    private static void TextProperties(XmlWriter w, string font, double size, bool bold, string color)
    {
        w.WriteStartElement("txPr", ChartNs);
        w.WriteElementString("bodyPr", DrawingNs, null);
        w.WriteElementString("lstStyle", DrawingNs, null);
        w.WriteStartElement("p", DrawingNs);
        w.WriteStartElement("pPr", DrawingNs);
        RunProperties(w, "defRPr", font, size, bold, color);
        w.WriteEndElement();
        w.WriteStartElement("endParaRPr", DrawingNs);
        w.WriteAttributeString("lang", "en-US");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void Line(XmlWriter w, string color, double width, string dash)
    {
        w.WriteStartElement("ln", DrawingNs);
        w.WriteAttributeString("w", Int((int)(width * 12700)));
        w.WriteAttributeString("cap", "rnd");
        Solid(w, color);
        w.WriteStartElement("prstDash", DrawingNs);
        w.WriteAttributeString("val", dash);
        w.WriteEndElement();
        w.WriteElementString("round", DrawingNs, null);
        w.WriteEndElement();
    }

    private static void Solid(XmlWriter w, string color)
    {
        w.WriteStartElement("solidFill", DrawingNs);
        w.WriteStartElement("srgbClr", DrawingNs);
        w.WriteAttributeString("val", XlsxWorkbook.Argb(color).Substring(2));
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void Val(XmlWriter w, string element, string value)
    {
        w.WriteStartElement(element, ChartNs);
        w.WriteAttributeString("val", value);
        w.WriteEndElement();
    }

    private static string Int(int value) => value.ToString(CultureInfo.InvariantCulture);

    private static string Num(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    #endregion
}
