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
            Anchor(w, image.Row, image.Column, image.Width, image.Height);
            w.WriteStartElement("pic", SheetDrawingNs);
            w.WriteStartElement("nvPicPr", SheetDrawingNs);
            w.WriteStartElement("cNvPr", SheetDrawingNs);
            w.WriteAttributeString("id", Int(charts.Count + j + 2));
            w.WriteAttributeString("name", "Picture " + Int(j + 1));
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
    public static void WriteChart(XmlWriter w, XlsxChart chart, string defaultFont)
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

        w.WriteStartElement(polar ? "radarChart" : "scatterChart", ChartNs);
        Val(w, polar ? "radarStyle" : "scatterStyle", polar ? "marker" : "smoothMarker");
        Val(w, "varyColors", "0");
        for (var i = 0; i < chart.Series.Count; i++)
        {
            Series(w, chart, chart.Series[i], i, polar);
        }

        Val(w, "axId", "1");
        Val(w, "axId", "2");
        w.WriteEndElement();
        if (polar)
        {
            AngleAxis(w, chart, font, look);
            Axis(w, chart.YAxis, 2, 1, "l", font, look, true);
        }
        else
        {
            Axis(w, chart.XAxis, 1, 2, "b", font, look, false);
            Axis(w, chart.YAxis, 2, 1, "l", font, look, false);
        }

        w.WriteEndElement();

        var legend = chart.ShowLegend ?? chart.Series.Count > 1;
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
    private static void Values(XmlWriter w, string name, XlsxSheet sheet, int column, int firstRow, int lastRow)
    {
        var col = "$" + XlsxWorkbook.ColumnName(column) + "$";
        w.WriteStartElement(name, ChartNs);
        w.WriteStartElement("numRef", ChartNs);
        w.WriteElementString("f", ChartNs, "'" + sheet.Name.Replace("'", "''") + "'!" + col + Int(firstRow) + ":" + col + Int(lastRow));
        w.WriteStartElement("numCache", ChartNs);
        w.WriteElementString("formatCode", ChartNs, "General");
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
    /// A value axis (both axes of a scatter chart are value axes). Each axis crosses the other at its minimum, so the
    /// axes stay in the bottom-left corner even when every value is negative.
    /// </summary>
    private static void Axis(XmlWriter w, XlsxAxis axis, int id, int crossId, string position, string font, XlsxChartLook look, bool radius)
    {
        w.WriteStartElement("valAx", ChartNs);
        Val(w, "axId", Int(id));
        w.WriteStartElement("scaling", ChartNs);
        Val(w, "orientation", "minMax");
        if (axis.Max != null) { Val(w, "max", Num(axis.Max.Value)); }
        if (axis.Min != null) { Val(w, "min", Num(axis.Min.Value)); }
        w.WriteEndElement();
        Val(w, "delete", "0");
        Val(w, "axPos", position);
        if (axis.Gridlines)
        {
            w.WriteStartElement("majorGridlines", ChartNs);
            w.WriteStartElement("spPr", ChartNs);
            Line(w, look.GridColor, 0.75, "solid");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        if (axis.Title != null)
        {
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

        w.WriteStartElement("numFmt", ChartNs);
        w.WriteAttributeString("formatCode", axis.NumberFormat ?? "General");
        w.WriteAttributeString("sourceLinked", axis.NumberFormat == null ? "1" : "0");
        w.WriteEndElement();
        Val(w, "majorTickMark", radius ? "cross" : "out");
        Val(w, "minorTickMark", "none");
        Val(w, "tickLblPos", radius ? "nextTo" : "low");
        w.WriteStartElement("spPr", ChartNs);
        Line(w, look.AxisColor, 0.75, "solid");
        w.WriteEndElement();
        TextProperties(w, font, look.TextSize, false, look.TextColor);
        Val(w, "crossAx", Int(crossId));
        Val(w, "crosses", radius ? "autoZero" : "min");
        Val(w, "crossBetween", radius ? "between" : "midCat");
        if (axis.MajorUnit != null) { Val(w, "majorUnit", Num(axis.MajorUnit.Value)); }
        w.WriteEndElement();
    }

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
