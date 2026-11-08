#nullable enable
namespace Divit.Docx;

using Divit.Packaging;
using System;
using System.Collections.Generic;
using System.Xml;

/// <content>
/// Tables and images.
/// </content>
public sealed partial class DocxDocument
{
    private const string FiguresStyle = "TableofFigures";
    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string WordDrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private const string PictureNs = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    /// <summary>
    /// The table style of a look; defined in the styles part.
    /// </summary>
    private static string TableStyleOf(DocxTableLook look) =>
        look == DocxTableLook.Grid ? "DivitGrid" : look == DocxTableLook.Plain ? "DivitPlain" : "DivitTable";

    /// <summary>
    /// A table: properties (style, width, layout, which conditional looks apply), the grid of column widths, then
    /// the rows. Spans write one cell for the columns they cover; a cell under a row span continues the merge.
    /// </summary>
    private void WriteTable(XmlWriter w, DocxTable table, Scope scope)
    {
        var state = StartTable(w, table, scope);
        foreach (var row in table.Rows)
        {
            WriteRow(w, row, state, scope);
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Opens a table with its properties and grid; the rows follow, and the caller closes it.
    /// </summary>
    private static TableState StartTable(XmlWriter w, DocxTable table, Scope scope)
    {
        var widths = new int[table.Columns];
        for (var c = 0; c < table.Columns; c++)
        {
            widths[c] = table.ColumnWidths != null ? DocxUnits.Twips(table.ColumnWidths[c]) : scope.Width / table.Columns;
        }

        var total = 0;
        foreach (var width in widths)
        {
            total += width;
        }

        w.WriteStartElement("tbl", Main);
        w.WriteStartElement("tblPr", Main);
        Val(w, "tblStyle", TableStyleOf(table.Look));
        Width(w, "tblW", total);
        if (table.Centered) { Val(w, "jc", "center"); }
        w.WriteStartElement("tblLayout", Main);
        w.WriteAttributeString("w", "type", Main, "fixed");
        w.WriteEndElement();
        w.WriteStartElement("tblLook", Main);
        w.WriteAttributeString("w", "val", Main, "04A0");
        w.WriteAttributeString("w", "firstRow", Main, "1");
        w.WriteAttributeString("w", "lastRow", Main, "0");
        w.WriteAttributeString("w", "firstColumn", Main, "0");
        w.WriteAttributeString("w", "lastColumn", Main, "0");
        w.WriteAttributeString("w", "noHBand", Main, "0");
        w.WriteAttributeString("w", "noVBand", Main, "1");
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("tblGrid", Main);
        foreach (var width in widths)
        {
            w.WriteStartElement("gridCol", Main);
            w.WriteAttributeString("w", "w", Main, Int(width));
            w.WriteEndElement();
        }

        w.WriteEndElement();
        return new TableState(table, widths);
    }

    /// <summary>
    /// One row: its properties, then a cell per column, or one for the columns a span covers; under a row span the
    /// cell continues the merge.
    /// </summary>
    private void WriteRow(XmlWriter w, DocxTableRow row, TableState state, Scope scope)
    {
        var table = state.Table;
        var widths = state.Widths;
        var covered = state.Covered;
        var coveredSpan = state.CoveredSpan;
        w.WriteStartElement("tr", Main);
        if (row.KeepTogether || row.Height != null || row.IsHeader)
        {
            w.WriteStartElement("trPr", Main);
            if (row.KeepTogether) { w.WriteElementString("cantSplit", Main, null); }
            if (row.Height is double height)
            {
                w.WriteStartElement("trHeight", Main);
                w.WriteAttributeString("w", "val", Main, Int(DocxUnits.Twips(height)));
                w.WriteAttributeString("w", "hRule", Main, "atLeast");
                w.WriteEndElement();
            }

            if (row.IsHeader) { w.WriteElementString("tblHeader", Main, null); }
            w.WriteEndElement();
        }

        var c = 0;
        while (c < table.Columns)
        {
            if (covered[c] > 0)
            {
                var span = coveredSpan[c];
                covered[c]--;
                CellStart(w, Sum(widths, c, span), span, "continue", null, DocxCellAlignment.Top);
                w.WriteElementString("p", Main, null);
                w.WriteEndElement();
                c += span;
                continue;
            }

            var cell = row.Cells[c];
            var columns = Math.Max(1, Math.Min(cell.ColumnSpan, table.Columns - c));
            string? merge = null;
            if (cell.RowSpan > 1)
            {
                merge = "restart";
                covered[c] = cell.RowSpan - 1;
                coveredSpan[c] = columns;
            }

            CellStart(w, Sum(widths, c, columns), columns, merge, cell.Shading, cell.VerticalAlignment);
            var inner = new Scope(scope.Parts, scope.Headings, scope.Captions) { Width = Sum(widths, c, columns) - 216 };
            WriteBlocks(w, cell, inner, table.AlignmentOf(c));
            if (cell.Blocks.Count == 0 || cell.Blocks[cell.Blocks.Count - 1] is DocxTable)
            {
                // A cell ends with a paragraph, also after a nested table.
                w.WriteElementString("p", Main, null);
            }

            w.WriteEndElement();
            c += columns;
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// What the rows of a table share while it is written: the column widths, and per column the rows still covered
    /// by a row span above and how many columns that span is wide.
    /// </summary>
    private sealed class TableState
    {
        public TableState(DocxTable table, int[] widths)
        {
            Table = table;
            Widths = widths;
            Covered = new int[table.Columns];
            CoveredSpan = new int[table.Columns];
        }

        public DocxTable Table { get; }

        public int[] Widths { get; }

        public int[] Covered { get; }

        public int[] CoveredSpan { get; }
    }

    /// <summary>
    /// Opens a cell with its properties in the order of the standard: width, span, merge, fill, alignment.
    /// </summary>
    private static void CellStart(XmlWriter w, int width, int span, string? merge, string? shading, DocxCellAlignment alignment)
    {
        w.WriteStartElement("tc", Main);
        w.WriteStartElement("tcPr", Main);
        Width(w, "tcW", width);
        if (span > 1) { Val(w, "gridSpan", Int(span)); }
        if (merge == "restart") { Val(w, "vMerge", "restart"); }
        else if (merge != null) { w.WriteElementString("vMerge", Main, null); }
        if (shading != null) { Shading(w, shading); }
        if (alignment != DocxCellAlignment.Top) { Val(w, "vAlign", alignment == DocxCellAlignment.Center ? "center" : "bottom"); }
        w.WriteEndElement();
    }

    private static void Width(XmlWriter w, string name, int twips)
    {
        w.WriteStartElement(name, Main);
        w.WriteAttributeString("w", "w", Main, Int(twips));
        w.WriteAttributeString("w", "type", Main, "dxa");
        w.WriteEndElement();
    }

    private static int Sum(int[] widths, int start, int count)
    {
        var sum = 0;
        for (var i = start; i < start + count; i++)
        {
            sum += widths[i];
        }

        return sum;
    }

    /// <summary>
    /// An image in the line (DrawingML inline picture), pointing at its media part through the part's relationship.
    /// </summary>
    private static void WriteImage(XmlWriter w, DocxRun run, Scope scope)
    {
        var id = ++scope.Parts.Drawings;
        var cx = run.ImageWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var cy = run.ImageHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        w.WriteStartElement("r", Main);
        WriteRunProperties(w, run.Font);
        w.WriteStartElement("drawing", Main);
        w.WriteStartElement("wp", "inline", WordDrawingNs);
        foreach (var side in new[] { "distT", "distB", "distL", "distR" })
        {
            w.WriteAttributeString(side, "0");
        }

        w.WriteStartElement("wp", "extent", WordDrawingNs);
        w.WriteAttributeString("cx", cx);
        w.WriteAttributeString("cy", cy);
        w.WriteEndElement();
        w.WriteStartElement("wp", "effectExtent", WordDrawingNs);
        foreach (var side in new[] { "l", "t", "r", "b" })
        {
            w.WriteAttributeString(side, "0");
        }

        w.WriteEndElement();
        w.WriteStartElement("wp", "docPr", WordDrawingNs);
        w.WriteAttributeString("id", Int(id));
        w.WriteAttributeString("name", "Picture " + Int(id));
        if (run.Description != null) { w.WriteAttributeString("descr", OfficeText.Clean(run.Description)); }
        w.WriteEndElement();
        w.WriteStartElement("wp", "cNvGraphicFramePr", WordDrawingNs);
        w.WriteStartElement("a", "graphicFrameLocks", DrawingNs);
        w.WriteAttributeString("noChangeAspect", "1");
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("a", "graphic", DrawingNs);
        w.WriteStartElement("a", "graphicData", DrawingNs);
        w.WriteAttributeString("uri", PictureNs);
        w.WriteStartElement("pic", "pic", PictureNs);
        w.WriteStartElement("pic", "nvPicPr", PictureNs);
        w.WriteStartElement("pic", "cNvPr", PictureNs);
        w.WriteAttributeString("id", "0");
        w.WriteAttributeString("name", "image" + Int(id) + "." + run.ImageInfo.Extension);
        w.WriteEndElement();
        w.WriteElementString("pic", "cNvPicPr", PictureNs, string.Empty);
        w.WriteEndElement();
        w.WriteStartElement("pic", "blipFill", PictureNs);
        w.WriteStartElement("a", "blip", DrawingNs);
        w.WriteAttributeString("r", "embed", RelNs, scope.Parts.ImageIds[run]);
        w.WriteEndElement();
        w.WriteStartElement("a", "stretch", DrawingNs);
        w.WriteElementString("a", "fillRect", DrawingNs, string.Empty);
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("pic", "spPr", PictureNs);
        w.WriteStartElement("a", "xfrm", DrawingNs);
        w.WriteStartElement("a", "off", DrawingNs);
        w.WriteAttributeString("x", "0");
        w.WriteAttributeString("y", "0");
        w.WriteEndElement();
        w.WriteStartElement("a", "ext", DrawingNs);
        w.WriteAttributeString("cx", cx);
        w.WriteAttributeString("cy", cy);
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("a", "prstGeom", DrawingNs);
        w.WriteAttributeString("prst", "rect");
        w.WriteElementString("a", "avLst", DrawingNs, string.Empty);
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// A chart in the line: an inline frame whose graphic is the chart part, through the part's relationship.
    /// </summary>
    private static void WriteChart(XmlWriter w, DocxRun run, Scope scope)
    {
        const string ChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var id = ++scope.Parts.Drawings;
        var cx = run.ImageWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var cy = run.ImageHeight.ToString(System.Globalization.CultureInfo.InvariantCulture);
        w.WriteStartElement("r", Main);
        w.WriteStartElement("drawing", Main);
        w.WriteStartElement("wp", "inline", WordDrawingNs);
        foreach (var side in new[] { "distT", "distB", "distL", "distR" })
        {
            w.WriteAttributeString(side, "0");
        }

        w.WriteStartElement("wp", "extent", WordDrawingNs);
        w.WriteAttributeString("cx", cx);
        w.WriteAttributeString("cy", cy);
        w.WriteEndElement();
        w.WriteStartElement("wp", "effectExtent", WordDrawingNs);
        foreach (var side in new[] { "l", "t", "r", "b" })
        {
            w.WriteAttributeString(side, "0");
        }

        w.WriteEndElement();
        w.WriteStartElement("wp", "docPr", WordDrawingNs);
        w.WriteAttributeString("id", Int(id));
        w.WriteAttributeString("name", "Chart " + Int(id));
        if (run.Chart!.Title != null) { w.WriteAttributeString("descr", OfficeText.Clean(run.Chart.Title)); }
        w.WriteEndElement();
        w.WriteElementString("wp", "cNvGraphicFramePr", WordDrawingNs, string.Empty);
        w.WriteStartElement("a", "graphic", DrawingNs);
        w.WriteStartElement("a", "graphicData", DrawingNs);
        w.WriteAttributeString("uri", ChartUri);
        w.WriteStartElement("c", "chart", ChartUri);
        w.WriteAttributeString("r", "id", RelNs, scope.Parts.ChartIds[run]);
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The captions of the document in order, as the lists of figures and tables show them.
    /// </summary>
    private List<(DocxCaptionKind Kind, string Text)> Captions()
    {
        var list = new List<(DocxCaptionKind, string)>();
        foreach (var section in _sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is DocxParagraph p && p.Caption is DocxCaptionKind kind)
                {
                    list.Add((kind, p.CaptionText ?? p.Text));
                }
            }
        }

        return list;
    }

    /// <summary>
    /// What the writers of one part share: the save's parts, the headings and captions for the contents, and the
    /// width the content may fill.
    /// </summary>
    private sealed class Scope
    {
        public Scope(Parts parts, IReadOnlyList<(int Level, string Number, string Text)> headings, IReadOnlyList<(DocxCaptionKind Kind, string Text)> captions)
        {
            Parts = parts;
            Headings = headings;
            Captions = captions;
        }

        public Parts Parts { get; }

        public IReadOnlyList<(int Level, string Number, string Text)> Headings { get; }

        public IReadOnlyList<(DocxCaptionKind Kind, string Text)> Captions { get; }

        /// <summary>
        /// Width between the margins (or of the cell) in twentieths of a point.
        /// </summary>
        public int Width { get; set; }
    }
}
