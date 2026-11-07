#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

/// <content>
/// The worksheet part and its companions: cells, rows, columns, view, print settings, comments and their VML
/// shapes.
/// </content>
public sealed partial class XlsxWorkbook
{
    #region Sheet part

    private void WriteSheet(XmlWriter w, XlsxSheet sheet, bool selected, SheetPlan plan, StyleTable styles, StringTable strings)
    {
        w.WriteStartElement("worksheet", MainNs);
        w.WriteAttributeString("xmlns", "r", null, RelNs);

        // sheetPr: tab colour, fit-to-page printing.
        w.WriteStartElement("sheetPr", MainNs);
        if (sheet.TabColor != null)
        {
            w.WriteStartElement("tabColor", MainNs);
            w.WriteAttributeString("rgb", Argb(sheet.TabColor));
            w.WriteEndElement();
        }

        w.WriteStartElement("pageSetUpPr", MainNs);
        w.WriteAttributeString("fitToPage", sheet.FitsToPage ? "1" : "0");
        w.WriteEndElement();
        w.WriteEndElement();

        // Cells in row order, then column order: the keys sort that way. Two parallel arrays sort faster than pairs.
        var count = sheet.Cells.Count;
        var keys = new long[count];
        var cells = new XlsxSheet.Cell[count];
        var n = 0;
        foreach (var pair in sheet.Cells)
        {
            keys[n] = pair.Key;
            cells[n] = pair.Value;
            n++;
        }

        Array.Sort(keys, cells);

        // Rows with a height, hidden or grouped are written even when they hold no cell.
        var heights = sheet.Rows.Keys.ToArray();
        Array.Sort(heights);

        // The used range of a streamed sheet is known only after its rows; the element is optional and left out.
        if (sheet.RowSource == null)
        {
            w.WriteStartElement("dimension", MainNs);
            if (count == 0)
            {
                w.WriteAttributeString("ref", "A1");
            }
            else
            {
                var lastColumn = 1;
                foreach (var key in keys)
                {
                    lastColumn = Math.Max(lastColumn, XlsxSheet.ColumnOf(key));
                }

                var lastRow = XlsxSheet.RowOf(keys[count - 1]);
                if (heights.Length > 0)
                {
                    lastRow = Math.Max(lastRow, heights[heights.Length - 1]);
                }

                w.WriteAttributeString("ref", "A1:" + CellName(lastRow, lastColumn));
            }

            w.WriteEndElement();
        }

        // sheetViews: gridlines, selected tab, frozen panes.
        w.WriteStartElement("sheetViews", MainNs);
        w.WriteStartElement("sheetView", MainNs);
        if (!sheet.ShowGridLines) { w.WriteAttributeString("showGridLines", "0"); }
        if (selected) { w.WriteAttributeString("tabSelected", "1"); }
        // A sheet without its own zoom opens at DefaultZoom; without zoomScaleNormal some Excel versions open at 100%.
        var zoom = sheet.Zoom >= 10 && sheet.Zoom <= 400 ? sheet.Zoom : DefaultZoom;
        if (zoom >= 10 && zoom <= 400)
        {
            w.WriteAttributeString("zoomScale", Int(zoom));
            w.WriteAttributeString("zoomScaleNormal", Int(zoom));
        }
        w.WriteAttributeString("workbookViewId", "0");
        if (sheet.FrozenRows > 0 || sheet.FrozenColumns > 0)
        {
            var pane = sheet.FrozenRows > 0 && sheet.FrozenColumns > 0 ? "bottomRight" : sheet.FrozenRows > 0 ? "bottomLeft" : "topRight";
            w.WriteStartElement("pane", MainNs);
            if (sheet.FrozenColumns > 0) { w.WriteAttributeString("xSplit", Int(sheet.FrozenColumns)); }
            if (sheet.FrozenRows > 0) { w.WriteAttributeString("ySplit", Int(sheet.FrozenRows)); }
            w.WriteAttributeString("topLeftCell", CellName(sheet.FrozenRows + 1, sheet.FrozenColumns + 1));
            w.WriteAttributeString("activePane", pane);
            w.WriteAttributeString("state", "frozen");
            w.WriteEndElement();
            w.WriteStartElement("selection", MainNs);
            w.WriteAttributeString("pane", pane);
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteEndElement();

        // Without customHeight Excel ignores defaultRowHeight and computes its own from the font (14.25 at 10 pt); chart
        // positions are computed from this height, so Excel must use the same one.
        w.WriteStartElement("sheetFormatPr", MainNs);
        w.WriteAttributeString("defaultRowHeight", Num(DefaultRowHeight));
        w.WriteAttributeString("customHeight", "1");

        // The deepest outline levels; Excel sizes the outline button bar from them.
        var rowLevels = sheet.Rows.Values.Select(l => l.Level).DefaultIfEmpty(0).Max();
        var columnLevels = sheet.Columns.Values.Select(l => l.Level).DefaultIfEmpty(0).Max();
        if (rowLevels > 0) { w.WriteAttributeString("outlineLevelRow", Int(rowLevels)); }
        if (columnLevels > 0) { w.WriteAttributeString("outlineLevelCol", Int(columnLevels)); }
        w.WriteEndElement();

        // cols: only columns with a width, hidden or grouped.
        if (sheet.Columns.Count > 0)
        {
            var numbers = sheet.Columns.Keys.ToArray();
            Array.Sort(numbers);
            w.WriteStartElement("cols", MainNs);
            foreach (var col in numbers)
            {
                var line = sheet.Columns[col];
                w.WriteStartElement("col", MainNs);
                w.WriteAttributeString("min", Int(col));
                w.WriteAttributeString("max", Int(col));
                w.WriteAttributeString("width", Num(line.Size ?? 8.43));
                if (line.Hidden) { w.WriteAttributeString("hidden", "1"); }
                w.WriteAttributeString("customWidth", "1");
                if (line.Level > 0) { w.WriteAttributeString("outlineLevel", Int(line.Level)); }
                if (line.Collapsed) { w.WriteAttributeString("collapsed", "1"); }
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        // sheetData: rows come from merging the sorted cell keys with the sorted row properties. Addresses and numbers
        // are formatted into one scratch buffer: a sheet holds up to millions of cells, and a string for each would
        // dominate the save.
        w.WriteStartElement("sheetData", MainNs);
        var buffer = new char[64];

        // A streamed sheet holds no cells: its callback writes them straight into this part.
        if (sheet.RowSource != null)
        {
            var rows = new XlsxRowWriter(new StreamSink(this, w, styles, strings, buffer, sheet.InlineStrings));
            sheet.RowSource(rows);
            rows.Finish();
        }

        var index = 0;
        var nextHeight = 0;
        while (index < count || nextHeight < heights.Length)
        {
            var cellRow = index < count ? XlsxSheet.RowOf(keys[index]) : int.MaxValue;
            var heightRow = nextHeight < heights.Length ? heights[nextHeight] : int.MaxValue;
            var row = Math.Min(cellRow, heightRow);

            w.WriteStartElement("row", MainNs);
            WriteNumberAttribute(w, "r", row, buffer);
            if (heightRow == row)
            {
                var line = sheet.Rows[row];
                if (line.Size is double height)
                {
                    w.WriteAttributeString("ht", Num(height));
                    w.WriteAttributeString("customHeight", "1");
                }

                if (line.Hidden) { w.WriteAttributeString("hidden", "1"); }
                if (line.Level > 0) { w.WriteAttributeString("outlineLevel", Int(line.Level)); }
                if (line.Collapsed) { w.WriteAttributeString("collapsed", "1"); }
                nextHeight++;
            }

            while (index < count && XlsxSheet.RowOf(keys[index]) == row)
            {
                _formulas |= cells[index].Formula != null;
                WriteCell(w, keys[index], cells[index], styles, strings, buffer);
                index++;
            }

            w.WriteEndElement();
        }

        w.WriteEndElement();

        // Protection (right after sheetData in the standard). Selecting cells stays allowed so links can be clicked.
        if (sheet.Protected)
        {
            w.WriteStartElement("sheetProtection", MainNs);
            if (sheet.Password != null && sheet.Password.Length > 0)
            {
                var salt = PasswordHash.Salt(Created == null ? null : "Divit|sheet|" + sheet.Name + "|" + Created.Value.Ticks.ToString(CultureInfo.InvariantCulture));
                w.WriteAttributeString("algorithmName", PasswordHash.Algorithm);
                w.WriteAttributeString("hashValue", PasswordHash.Sha512(sheet.Password, salt));
                w.WriteAttributeString("saltValue", Convert.ToBase64String(salt));
                w.WriteAttributeString("spinCount", Int(PasswordHash.SpinCount));
            }

            w.WriteAttributeString("sheet", "1");
            w.WriteAttributeString("objects", "1");
            w.WriteAttributeString("scenarios", "1");

            // These attributes say what is locked: "0" allows the action.
            var allowed = sheet.Allowed;
            foreach (var (flag, name) in AllowAttributes)
            {
                if ((allowed & flag) != 0) { w.WriteAttributeString(name, "0"); }
            }

            w.WriteEndElement();
        }

        if (sheet.AutoFilterRange is var range && range != null)
        {
            w.WriteStartElement("autoFilter", MainNs);
            w.WriteAttributeString("ref", CellName(range.Value.Row1, range.Value.Col1) + ":" + CellName(range.Value.Row2, range.Value.Col2));
            w.WriteEndElement();
        }

        if (sheet.Merges.Count > 0)
        {
            w.WriteStartElement("mergeCells", MainNs);
            w.WriteAttributeString("count", Int(sheet.Merges.Count));
            foreach (var m in sheet.Merges)
            {
                w.WriteStartElement("mergeCell", MainNs);
                w.WriteAttributeString("ref", CellName(m.Row1, m.Col1) + ":" + CellName(m.Row2, m.Col2));
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        // Conditional formats, one block per range; the order they were added is their priority.
        var priority = 0;
        foreach (var conditional in sheet.Conditionals)
        {
            w.WriteStartElement("conditionalFormatting", MainNs);
            w.WriteAttributeString("sqref", RangeName(conditional.Row1, conditional.Column1, conditional.Row2, conditional.Column2));
            w.WriteStartElement("cfRule", MainNs);
            switch (conditional)
            {
                case XlsxSheet.XlsxColorScale scale:
                    WriteColorScale(w, scale, ++priority);
                    break;
                case XlsxSheet.XlsxDataBar bar:
                    WriteDataBar(w, bar, ++priority);
                    break;
                case XlsxSheet.XlsxIcons icons:
                    WriteIconSet(w, icons, ++priority);
                    break;
                case XlsxSheet.XlsxHighlight highlight:
                    WriteHighlight(w, highlight, ++priority, styles);
                    break;
            }

            w.WriteEndElement();
            w.WriteEndElement();
        }

        if (sheet.Validations.Count > 0)
        {
            w.WriteStartElement("dataValidations", MainNs);
            w.WriteAttributeString("count", Int(sheet.Validations.Count));
            foreach (var item in sheet.Validations)
            {
                WriteValidation(w, item.Validation, RangeName(item.Row1, item.Col1, item.Row2, item.Col2));
            }

            w.WriteEndElement();
        }

        // A link inside the workbook needs no relationship: location is enough. A link outside names its relationship,
        // which holds the address.
        if (sheet.Links.Count > 0 || sheet.UrlLinks.Count > 0)
        {
            w.WriteStartElement("hyperlinks", MainNs);
            foreach (var link in sheet.Links)
            {
                w.WriteStartElement("hyperlink", MainNs);
                w.WriteAttributeString("ref", CellName(link.Row, link.Col));
                w.WriteAttributeString("location", "'" + link.Sheet.Replace("'", "''") + "'!A1");
                w.WriteAttributeString("display", link.Sheet);
                w.WriteEndElement();
            }

            for (var i = 0; i < sheet.UrlLinks.Count; i++)
            {
                var link = sheet.UrlLinks[i];
                w.WriteStartElement("hyperlink", MainNs);
                w.WriteAttributeString("ref", CellName(link.Row, link.Col));
                w.WriteAttributeString("id", RelNs, plan.LinkIds[i]);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        if (sheet.PrintCenteredHorizontally || sheet.PrintCenteredVertically || sheet.PrintGridLines)
        {
            w.WriteStartElement("printOptions", MainNs);
            if (sheet.PrintCenteredHorizontally) { w.WriteAttributeString("horizontalCentered", "1"); }
            if (sheet.PrintCenteredVertically) { w.WriteAttributeString("verticalCentered", "1"); }
            if (sheet.PrintGridLines) { w.WriteAttributeString("gridLines", "1"); }
            w.WriteEndElement();
        }

        var margins = sheet.Margins;
        w.WriteStartElement("pageMargins", MainNs);
        w.WriteAttributeString("left", Inches(margins.Left));
        w.WriteAttributeString("right", Inches(margins.Right));
        w.WriteAttributeString("top", Inches(margins.Top));
        w.WriteAttributeString("bottom", Inches(margins.Bottom));
        w.WriteAttributeString("header", Inches(margins.Header));
        w.WriteAttributeString("footer", Inches(margins.Footer));
        w.WriteEndElement();

        w.WriteStartElement("pageSetup", MainNs);
        w.WriteAttributeString("paperSize", Int((int)sheet.PaperSize));
        w.WriteAttributeString("orientation", sheet.Landscape ? "landscape" : "portrait");
        w.WriteAttributeString("fitToWidth", Int(sheet.FitWidth));
        w.WriteAttributeString("fitToHeight", Int(sheet.FitHeight));
        w.WriteEndElement();

        // Header and footer: sections &L &C &R, &P page, &N page count.
        var even = sheet.EvenPageHeader != null || sheet.EvenPageFooter != null;
        var first = sheet.FirstPageHeader != null || sheet.FirstPageFooter != null;
        if (sheet.Header != null || sheet.Footer != null || even || first)
        {
            // Odd pages (all pages unless even pages differ), then even, then first, in the order of the standard.
            w.WriteStartElement("headerFooter", MainNs);
            if (even) { w.WriteAttributeString("differentOddEven", "1"); }
            if (first) { w.WriteAttributeString("differentFirst", "1"); }
            HeaderPart(w, "oddHeader", sheet.Header);
            HeaderPart(w, "oddFooter", sheet.Footer);
            HeaderPart(w, "evenHeader", sheet.EvenPageHeader);
            HeaderPart(w, "evenFooter", sheet.EvenPageFooter);
            HeaderPart(w, "firstHeader", sheet.FirstPageHeader);
            HeaderPart(w, "firstFooter", sheet.FirstPageFooter);
            w.WriteEndElement();
        }

        // Manual page breaks: id is the last row (column) of a page; max spans the whole other direction.
        Breaks(w, "rowBreaks", sheet.RowBreaks, XlsxSheet.MaxColumn - 1);
        Breaks(w, "colBreaks", sheet.ColumnBreaks, XlsxSheet.MaxRow - 1);

        // Ignored warnings: after headerFooter and before drawing, in the order of the standard.
        if (sheet.IgnoreNumbersAsText)
        {
            w.WriteStartElement("ignoredErrors", MainNs);
            w.WriteStartElement("ignoredError", MainNs);
            w.WriteAttributeString("sqref", "A1:XFD1048576");
            w.WriteAttributeString("numberStoredAsText", "1");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        // Charts and images live in the drawing layer; notes in the legacy VML drawing right after it.
        if (plan.DrawingId != null)
        {
            w.WriteStartElement("drawing", MainNs);
            w.WriteAttributeString("id", RelNs, plan.DrawingId);
            w.WriteEndElement();
        }

        if (plan.NoteShapesId != null)
        {
            w.WriteStartElement("legacyDrawing", MainNs);
            w.WriteAttributeString("id", RelNs, plan.NoteShapesId);
            w.WriteEndElement();
        }

        // Tables come last in the worksheet; pivot tables are listed in the sheet's relationships only.
        if (plan.TableIds.Count > 0)
        {
            w.WriteStartElement("tableParts", MainNs);
            w.WriteAttributeString("count", Int(plan.TableIds.Count));
            foreach (var id in plan.TableIds)
            {
                w.WriteStartElement("tablePart", MainNs);
                w.WriteAttributeString("id", RelNs, id);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private static void HeaderPart(XmlWriter w, string name, string? text)
    {
        if (text != null)
        {
            w.WriteElementString(name, MainNs, OfficeText.Clean(text));
        }
    }

    private static void Breaks(XmlWriter w, string name, IReadOnlyCollection<int> breaks, int max)
    {
        if (breaks.Count == 0)
        {
            return;
        }

        w.WriteStartElement(name, MainNs);
        w.WriteAttributeString("count", Int(breaks.Count));
        w.WriteAttributeString("manualBreakCount", Int(breaks.Count));
        foreach (var at in breaks)
        {
            w.WriteStartElement("brk", MainNs);
            w.WriteAttributeString("id", Int(at));
            w.WriteAttributeString("max", Int(max));
            w.WriteAttributeString("man", "1");
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// The sheetProtection attribute of each permission.
    /// </summary>
    private static readonly (XlsxAllow Flag, string Name)[] AllowAttributes =
    {
        (XlsxAllow.FormatCells, "formatCells"),
        (XlsxAllow.FormatColumns, "formatColumns"),
        (XlsxAllow.FormatRows, "formatRows"),
        (XlsxAllow.InsertColumns, "insertColumns"),
        (XlsxAllow.InsertRows, "insertRows"),
        (XlsxAllow.InsertHyperlinks, "insertHyperlinks"),
        (XlsxAllow.DeleteColumns, "deleteColumns"),
        (XlsxAllow.DeleteRows, "deleteRows"),
        (XlsxAllow.Sort, "sort"),
        (XlsxAllow.AutoFilter, "autoFilter"),
        (XlsxAllow.PivotTables, "pivotTables"),
    };

    /// <summary>
    /// A range address; one cell when both corners are the same.
    /// </summary>
    private static string RangeName(int row1, int column1, int row2, int column2) =>
        row1 == row2 && column1 == column2 ? CellName(row1, column1) : CellName(row1, column1) + ":" + CellName(row2, column2);

    /// <summary>
    /// Colours for the lowest / [median] / highest value.
    /// </summary>
    private static void WriteColorScale(XmlWriter w, XlsxSheet.XlsxColorScale scale, int priority)
    {
        w.WriteAttributeString("type", "colorScale");
        w.WriteAttributeString("priority", Int(priority));
        w.WriteStartElement("colorScale", MainNs);
        Cfvo(w, "min", null);
        if (scale.Mid != null) { Cfvo(w, "percentile", "50"); }
        Cfvo(w, "max", null);
        foreach (var color in scale.Mid != null ? new[] { scale.Low, scale.Mid, scale.High } : new[] { scale.Low, scale.High })
        {
            w.WriteStartElement("color", MainNs);
            w.WriteAttributeString("rgb", Argb(color));
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// A bar from the lowest to the highest value of the range, in the Excel 2007 form every reader knows.
    /// </summary>
    private static void WriteDataBar(XmlWriter w, XlsxSheet.XlsxDataBar bar, int priority)
    {
        w.WriteAttributeString("type", "dataBar");
        w.WriteAttributeString("priority", Int(priority));
        w.WriteStartElement("dataBar", MainNs);
        Cfvo(w, "min", null);
        Cfvo(w, "max", null);
        w.WriteStartElement("color", MainNs);
        w.WriteAttributeString("rgb", Argb(bar.Color));
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// Icons by percentage of the value span, in equal parts: 0, 33, 67 for three icons.
    /// </summary>
    private static void WriteIconSet(XmlWriter w, XlsxSheet.XlsxIcons icons, int priority)
    {
        var (name, count) = IconSetName(icons.Set);
        w.WriteAttributeString("type", "iconSet");
        w.WriteAttributeString("priority", Int(priority));
        w.WriteStartElement("iconSet", MainNs);
        w.WriteAttributeString("iconSet", name);
        if (!icons.ShowValue) { w.WriteAttributeString("showValue", "0"); }
        if (icons.Reverse) { w.WriteAttributeString("reverse", "1"); }
        for (var i = 0; i < count; i++)
        {
            Cfvo(w, "percent", Int((int)Math.Round(100.0 * i / count, MidpointRounding.AwayFromZero)));
        }

        w.WriteEndElement();
    }

    private static (string Name, int Count) IconSetName(XlsxIconSet set)
    {
        switch (set)
        {
            case XlsxIconSet.ThreeTrafficLightsFramed: return ("3TrafficLights2", 3);
            case XlsxIconSet.ThreeSigns: return ("3Signs", 3);
            case XlsxIconSet.ThreeArrows: return ("3Arrows", 3);
            case XlsxIconSet.ThreeArrowsGray: return ("3ArrowsGray", 3);
            case XlsxIconSet.ThreeFlags: return ("3Flags", 3);
            case XlsxIconSet.ThreeSymbols: return ("3Symbols", 3);
            case XlsxIconSet.ThreeSymbolsPlain: return ("3Symbols2", 3);
            case XlsxIconSet.FourArrows: return ("4Arrows", 4);
            case XlsxIconSet.FourArrowsGray: return ("4ArrowsGray", 4);
            case XlsxIconSet.FourRedToBlack: return ("4RedToBlack", 4);
            case XlsxIconSet.FourRating: return ("4Rating", 4);
            case XlsxIconSet.FourTrafficLights: return ("4TrafficLights", 4);
            case XlsxIconSet.FiveArrows: return ("5Arrows", 5);
            case XlsxIconSet.FiveArrowsGray: return ("5ArrowsGray", 5);
            case XlsxIconSet.FiveRating: return ("5Rating", 5);
            case XlsxIconSet.FiveQuarters: return ("5Quarters", 5);
            default: return ("3TrafficLights1", 3);
        }
    }

    /// <summary>
    /// A rule with a differential format: only what the look sets changes on matching cells.
    /// </summary>
    private static void WriteHighlight(XmlWriter w, XlsxSheet.XlsxHighlight highlight, int priority, StyleTable styles)
    {
        var rule = highlight.Rule;
        w.WriteAttributeString("type", rule.Type);
        w.WriteAttributeString("dxfId", Int(styles.Differential(highlight.Look)));
        w.WriteAttributeString("priority", Int(priority));
        if (rule.Type == "top10")
        {
            w.WriteAttributeString("rank", Int(rule.Rank));
            if (rule.Percent) { w.WriteAttributeString("percent", "1"); }
            if (rule.Bottom) { w.WriteAttributeString("bottom", "1"); }
        }

        if (rule.Below) { w.WriteAttributeString("aboveAverage", "0"); }
        if (rule.Operator != null) { w.WriteAttributeString("operator", rule.Operator); }
        if (rule.Text != null) { w.WriteAttributeString("text", OfficeText.Clean(rule.Text)); }

        // Text and blank rules store the formula Excel evaluates, relative to the top-left cell.
        var formula = rule.CellFormula(CellName(highlight.Row1, highlight.Column1));
        foreach (var text in formula != null ? new[] { formula } : rule.Formulas.ToArray())
        {
            w.WriteElementString("formula", MainNs, OfficeText.Clean(text));
        }
    }

    private static void WriteValidation(XmlWriter w, XlsxValidation validation, string range)
    {
        validation.CheckLengths();
        w.WriteStartElement("dataValidation", MainNs);
        w.WriteAttributeString("type", validation.Type);
        if (validation.ErrorStyle == XlsxErrorStyle.Warning) { w.WriteAttributeString("errorStyle", "warning"); }
        if (validation.ErrorStyle == XlsxErrorStyle.Information) { w.WriteAttributeString("errorStyle", "information"); }
        if (validation.Operator != null) { w.WriteAttributeString("operator", validation.Operator); }
        if (validation.AllowBlank) { w.WriteAttributeString("allowBlank", "1"); }

        // showDropDown="1" hides the arrow: the attribute is named the wrong way round in the standard.
        if (validation.Type == "list" && !validation.ShowDropDown) { w.WriteAttributeString("showDropDown", "1"); }
        w.WriteAttributeString("showInputMessage", "1");
        w.WriteAttributeString("showErrorMessage", "1");
        if (validation.ErrorTitle != null) { w.WriteAttributeString("errorTitle", OfficeText.Clean(validation.ErrorTitle)); }
        if (validation.ErrorMessage != null) { w.WriteAttributeString("error", OfficeText.Clean(validation.ErrorMessage)); }
        if (validation.InputTitle != null) { w.WriteAttributeString("promptTitle", OfficeText.Clean(validation.InputTitle)); }
        if (validation.InputMessage != null) { w.WriteAttributeString("prompt", OfficeText.Clean(validation.InputMessage)); }
        w.WriteAttributeString("sqref", range);
        w.WriteElementString("formula1", MainNs, OfficeText.Clean(validation.Formula1));
        if (validation.Formula2 != null) { w.WriteElementString("formula2", MainNs, validation.Formula2); }
        w.WriteEndElement();
    }

    /// <summary>
    /// The comment list of a sheet: authors once each, then the notes in cell order.
    /// </summary>
    private void WriteComments(XmlWriter w, XlsxSheet sheet)
    {
        var authors = new List<string>();
        foreach (var comment in sheet.Comments)
        {
            var author = comment.Author ?? Author ?? Application;
            if (!authors.Contains(author))
            {
                authors.Add(author);
            }
        }

        w.WriteStartElement("comments", MainNs);
        w.WriteStartElement("authors", MainNs);
        foreach (var author in authors)
        {
            w.WriteElementString("author", MainNs, OfficeText.Clean(author));
        }

        w.WriteEndElement();
        w.WriteStartElement("commentList", MainNs);
        foreach (var comment in sheet.Comments)
        {
            w.WriteStartElement("comment", MainNs);
            w.WriteAttributeString("ref", CellName(comment.Row, comment.Column));
            w.WriteAttributeString("authorId", Int(authors.IndexOf(comment.Author ?? Author ?? Application)));
            w.WriteStartElement("text", MainNs);
            if (comment.Rich != null)
            {
                // Pieces with fonts, written as in the shared strings.
                new StringTable(this).Runs(w, comment.Rich);
            }
            else
            {
                w.WriteStartElement("t", MainNs);
                w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(OfficeText.Clean(comment.Text));
                w.WriteEndElement();
            }

            w.WriteEndElement();
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The VML shapes Excel draws notes with: one hidden yellow box per note, anchored next to its cell. Excel shows no
    /// note without its shape, although the text lives in the comment list.
    /// </summary>
    /// <param name="w">The writer.</param>
    /// <param name="sheet">The sheet.</param>
    /// <param name="block">Number of the VML part; shape ids are counted in blocks of 1024 per part.</param>
    private static void WriteNoteShapes(XmlWriter w, XlsxSheet sheet, int block)
    {
        const string V = "urn:schemas-microsoft-com:vml";
        const string O = "urn:schemas-microsoft-com:office:office";
        const string X = "urn:schemas-microsoft-com:office:excel";
        w.WriteStartElement("xml");
        w.WriteAttributeString("xmlns", "v", null, V);
        w.WriteAttributeString("xmlns", "o", null, O);
        w.WriteAttributeString("xmlns", "x", null, X);

        w.WriteStartElement("o", "shapelayout", O);
        w.WriteAttributeString("v", "ext", V, "edit");
        w.WriteStartElement("o", "idmap", O);
        w.WriteAttributeString("v", "ext", V, "edit");
        w.WriteAttributeString("data", Int(block));
        w.WriteEndElement();
        w.WriteEndElement();

        // Shape type 202 is a text box.
        w.WriteStartElement("v", "shapetype", V);
        w.WriteAttributeString("id", "_x0000_t202");
        w.WriteAttributeString("coordsize", "21600,21600");
        w.WriteAttributeString("o", "spt", O, "202");
        w.WriteAttributeString("path", "m,l,21600r21600,l21600,xe");
        w.WriteStartElement("v", "stroke", V);
        w.WriteAttributeString("joinstyle", "miter");
        w.WriteEndElement();
        w.WriteStartElement("v", "path", V);
        w.WriteAttributeString("gradientshapeok", "t");
        w.WriteAttributeString("o", "connecttype", O, "rect");
        w.WriteEndElement();
        w.WriteEndElement();

        var shape = block * 1024;
        foreach (var comment in sheet.Comments)
        {
            w.WriteStartElement("v", "shape", V);
            w.WriteAttributeString("id", "_x0000_s" + Int(++shape));
            w.WriteAttributeString("type", "#_x0000_t202");
            w.WriteAttributeString("style", "position:absolute;margin-left:60pt;margin-top:2pt;width:144pt;height:72pt;z-index:" + Int(shape - block * 1024) + ";visibility:hidden");
            w.WriteAttributeString("fillcolor", "#ffffe1");
            w.WriteAttributeString("o", "insetmode", O, "auto");
            w.WriteStartElement("v", "fill", V);
            w.WriteAttributeString("color2", "#ffffe1");
            w.WriteEndElement();
            w.WriteStartElement("v", "shadow", V);
            w.WriteAttributeString("on", "t");
            w.WriteAttributeString("color", "black");
            w.WriteAttributeString("obscured", "t");
            w.WriteEndElement();
            w.WriteStartElement("v", "path", V);
            w.WriteAttributeString("o", "connecttype", O, "none");
            w.WriteEndElement();
            w.WriteStartElement("v", "textbox", V);
            w.WriteAttributeString("style", "mso-direction-alt:auto");
            w.WriteStartElement("div");
            w.WriteAttributeString("style", "text-align:left");
            w.WriteEndElement();
            w.WriteEndElement();

            // Anchor: column, x offset, row, y offset of the top-left and bottom-right corners, zero-based. The box
            // starts one column right of the cell and spans two columns and four rows.
            var row = comment.Row - 1;
            var column = comment.Column - 1;
            w.WriteStartElement("x", "ClientData", X);
            w.WriteAttributeString("ObjectType", "Note");
            w.WriteElementString("x", "MoveWithCells", X, string.Empty);
            w.WriteElementString("x", "SizeWithCells", X, string.Empty);
            w.WriteElementString("x", "Anchor", X,
                Int(column + 1) + ", 15, " + Int(Math.Max(0, row - 1)) + ", 10, " + Int(column + 1 + comment.Columns) + ", 15, " + Int(Math.Max(0, row - 1) + comment.Rows) + ", 4");
            w.WriteElementString("x", "AutoFill", X, "False");
            w.WriteElementString("x", "Row", X, Int(row));
            w.WriteElementString("x", "Column", X, Int(column));
            w.WriteEndElement();
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private static void WriteCell(XmlWriter w, long key, XlsxSheet.Cell cell, StyleTable styles, StringTable strings, char[] buffer)
    {
        var style = styles.Index(cell.Style);
        if (cell.Value == null && cell.Formula == null && style == 0)
        {
            return;
        }

        w.WriteStartElement("c", MainNs);
        w.WriteStartAttribute("r");
        w.WriteChars(buffer, 0, FormatCellName(buffer, XlsxSheet.RowOf(key), XlsxSheet.ColumnOf(key)));
        w.WriteEndAttribute();
        if (style != 0)
        {
            WriteNumberAttribute(w, "s", style, buffer);
        }

        // A formula keeps its result inline (t="str" for text) instead of in the shared strings; without a result the
        // cell holds the formula alone and Excel computes it on opening.
        if (cell.Formula != null)
        {
            if (cell.Value is string) { w.WriteAttributeString("t", "str"); }
            if (cell.Value is bool) { w.WriteAttributeString("t", "b"); }
            w.WriteElementString("f", MainNs, OfficeText.Clean(cell.Formula));
            switch (cell.Value)
            {
                case string result:
                    w.WriteElementString("v", MainNs, OfficeText.Clean(result));
                    break;
                case bool result:
                    w.WriteElementString("v", MainNs, result ? "1" : "0");
                    break;
                case double result:
                    w.WriteStartElement("v", MainNs);
                    w.WriteChars(buffer, 0, OfficeText.FormatDouble(buffer, result));
                    w.WriteEndElement();
                    break;
            }
        }
        else if (cell.Value is string text)
        {
            w.WriteAttributeString("t", "s");
            w.WriteStartElement("v", MainNs);
            w.WriteChars(buffer, 0, OfficeText.FormatInt(buffer, 0, strings.Index(text)));
            w.WriteEndElement();
        }
        else if (cell.Value is XlsxRichText rich)
        {
            w.WriteAttributeString("t", "s");
            w.WriteStartElement("v", MainNs);
            w.WriteChars(buffer, 0, OfficeText.FormatInt(buffer, 0, strings.Index(rich)));
            w.WriteEndElement();
        }
        else if (cell.Value is bool flag)
        {
            w.WriteAttributeString("t", "b");
            w.WriteElementString("v", MainNs, flag ? "1" : "0");
        }
        else if (cell.Value is double number)
        {
            w.WriteStartElement("v", MainNs);
            w.WriteChars(buffer, 0, OfficeText.FormatDouble(buffer, number));
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private static void WriteNumberAttribute(XmlWriter w, string name, int value, char[] buffer)
    {
        w.WriteStartAttribute(name);
        w.WriteChars(buffer, 0, OfficeText.FormatInt(buffer, 0, value));
        w.WriteEndAttribute();
    }

    /// <summary>
    /// Writes a cell address (B3) into the buffer; returns its length.
    /// </summary>
    private static int FormatCellName(char[] buffer, int row, int column)
    {
        // At most three letters (XFD); built backwards, then moved to the front.
        var start = 3;
        while (column > 0)
        {
            buffer[--start] = (char)('A' + ((column - 1) % 26));
            column = (column - 1) / 26;
        }

        var letters = 3 - start;
        for (var i = 0; i < letters; i++)
        {
            buffer[i] = buffer[start + i];
        }

        return OfficeText.FormatInt(buffer, letters, row);
    }

    #endregion
}
