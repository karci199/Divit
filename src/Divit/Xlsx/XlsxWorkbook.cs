#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

/// <summary>
/// An Excel workbook (.xlsx): collects sheets and writes them to disk with <see cref="Save"/>.
/// </summary>
/// <remarks>
/// No external library. An xlsx file is a zip of XML parts (ECMA-376); .NET's own
/// <c>System.IO.Compression</c> and <c>XmlWriter</c> are enough. Parts and the elements inside them are written in the
/// order the standard requires: Excel offers to "repair" a file in the wrong order, OnlyOffice and LibreOffice misread it silently.
/// <para>Text goes into the shared string table (sharedStrings), so each distinct text is stored once. Styles are made
/// unique the same way: a thousand cells that look alike point at one style record.</para>
/// <para>The file is first written next to the target as <c>.tmp</c> and moved into place when complete, so an
/// interrupted save never leaves a half-written file.</para>
/// </remarks>
public sealed class XlsxWorkbook
{
    #region Fields

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";

    private readonly List<XlsxSheet> _sheets = new List<XlsxSheet>();

    /// <summary>
    /// Number of drawing layers and charts written by the last save; the content type list is built from them.
    /// </summary>
    private int _drawings;

    private int _charts;

    #endregion

    #region Properties

    /// <summary>
    /// Default font of the workbook; cells whose style names no font use it. The column width unit is also
    /// computed from it.
    /// </summary>
    public string DefaultFontName { get; set; } = "Segoe UI";

    /// <summary>
    /// Default font size in points.
    /// </summary>
    public double DefaultFontSize { get; set; } = 10;

    /// <summary>
    /// Height (pt) of a row without an explicit height: font size × 1.5. Excel uses the same value (customHeight).
    /// </summary>
    public double DefaultRowHeight => DefaultFontSize * 1.5;

    /// <summary>
    /// Application name written into the file properties.
    /// </summary>
    public string Application { get; set; } = "Divit";

    /// <summary>
    /// Opening zoom of sheets without their own zoom, percent (10–400).
    /// </summary>
    public int DefaultZoom { get; set; } = 100;

    /// <summary>
    /// The sheets in tab order.
    /// </summary>
    public IReadOnlyList<XlsxSheet> Sheets => _sheets;

    /// <summary>
    /// File property: title. Shown in Windows Explorer and on Excel's File → Info screen, and searchable.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// File property: subject.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// File property: keywords (for example a serial number, so a search in Explorer finds the file).
    /// </summary>
    public string? Keywords { get; set; }

    /// <summary>
    /// File property: description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// File property: author; <see cref="Application"/> when empty.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// File property: creation and modification time (UTC); the time of saving when empty. Set it to get byte-identical
    /// files from identical input.
    /// </summary>
    public DateTime? Created { get; set; }

    #endregion

    #region Public methods

    /// <summary>
    /// Adds a new sheet at the end.
    /// </summary>
    /// <param name="name">Tab name: 1–31 characters, no <c>: \ / ? * [ ]</c>, unique in the workbook (ignoring case).</param>
    /// <returns>The new sheet.</returns>
    public XlsxSheet AddSheet(string name)
    {
        var problem = SheetNameProblem(name);
        if (problem != null)
        {
            throw new ArgumentException(problem, nameof(name));
        }

        var sheet = new XlsxSheet(name);
        _sheets.Add(sheet);
        return sheet;
    }

    /// <summary>
    /// Whether a name can be a sheet name; the reason when it cannot, <c>null</c> when it can.
    /// </summary>
    public string? SheetNameProblem(string name)
    {
        if (name.Length == 0 || name.Length > 31)
        {
            return Messages.Format(Strings.SheetNameLength, name);
        }

        if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0 || name[0] == '\'' || name[name.Length - 1] == '\'')
        {
            return Messages.Format(Strings.SheetNameCharacters, name);
        }

        foreach (var sheet in _sheets)
        {
            if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return Messages.Format(Strings.SheetNameDuplicate, name);
            }
        }

        return null;
    }

    /// <summary>
    /// Writes the workbook; an existing target is overwritten.
    /// </summary>
    /// <param name="path">Target file.</param>
    public void Save(string path)
    {
        if (_sheets.Count == 0)
        {
            throw new InvalidOperationException(Messages.Get(Strings.WorkbookEmpty));
        }

        var styles = new StyleTable(this);
        var strings = new StringTable();
        var temp = path + ".tmp";

        try
        {
            using (var file = new FileStream(temp, FileMode.Create, FileAccess.Write, FileShare.None, 65536))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                // Sheets first: the style and string tables fill while they are written.
                for (var i = 0; i < _sheets.Count; i++)
                {
                    Part(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", w => WriteSheet(w, _sheets[i], i == 0, styles, strings));
                }

                // Every sheet with charts or images: sheet → drawing layer → charts and images. Numbers run across the workbook.
                _drawings = 0;
                _charts = 0;
                var media = 0;
                for (var i = 0; i < _sheets.Count; i++)
                {
                    var charts = _sheets[i].Charts;
                    var images = _sheets[i].Images;
                    if (charts.Count == 0 && images.Count == 0)
                    {
                        continue;
                    }

                    var drawing = ++_drawings;
                    var firstChart = _charts + 1;
                    var firstImage = media + 1;
                    foreach (var image in images)
                    {
                        var entry = Entry(zip, "xl/media/image" + (++media) + ".png", CompressionLevel.NoCompression);
                        using (var stream = entry.Open())
                        {
                            stream.Write(image.Png, 0, image.Png.Length);
                        }
                    }

                    Part(zip, "xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels", w =>
                    {
                        w.WriteStartElement("Relationships", PackageRelNs);
                        Relationship(w, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "../drawings/drawing" + drawing + ".xml");
                        w.WriteEndElement();
                    });
                    Part(zip, "xl/drawings/drawing" + drawing + ".xml", w => XlsxChartWriter.WriteDrawing(w, charts, images));
                    Part(zip, "xl/drawings/_rels/drawing" + drawing + ".xml.rels", w =>
                    {
                        w.WriteStartElement("Relationships", PackageRelNs);
                        for (var k = 0; k < charts.Count; k++)
                        {
                            Relationship(w, "rId" + (k + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart" + (firstChart + k) + ".xml");
                        }

                        for (var k = 0; k < images.Count; k++)
                        {
                            Relationship(w, "rId" + (charts.Count + k + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "../media/image" + (firstImage + k) + ".png");
                        }

                        w.WriteEndElement();
                    });
                    foreach (var chart in charts)
                    {
                        var number = ++_charts;
                        Part(zip, "xl/charts/chart" + number + ".xml", w => XlsxChartWriter.WriteChart(w, chart, DefaultFontName));
                    }
                }

                Part(zip, "xl/styles.xml", styles.Write);
                Part(zip, "xl/sharedStrings.xml", strings.Write);
                Part(zip, "xl/workbook.xml", WriteWorkbook);
                Part(zip, "xl/_rels/workbook.xml.rels", WriteWorkbookRels);
                Part(zip, "docProps/core.xml", WriteCore);
                Part(zip, "docProps/app.xml", WriteApp);
                Part(zip, "_rels/.rels", WriteRootRels);
                Part(zip, "[Content_Types].xml", WriteContentTypes);
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            File.Move(temp, path);
        }
        finally
        {
            if (File.Exists(temp))
            {
                File.Delete(temp);
            }
        }
    }

    /// <summary>
    /// Letters of a column number: 1 → A, 27 → AA.
    /// </summary>
    public static string ColumnName(int column)
    {
        var name = string.Empty;
        while (column > 0)
        {
            var rest = (column - 1) % 26;
            name = (char)('A' + rest) + name;
            column = (column - 1) / 26;
        }

        return name;
    }

    /// <summary>
    /// Cell address: (3, 2) → B3.
    /// </summary>
    public static string CellName(int row, int column) => ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

    #endregion

    #region Parts

    private delegate void PartWriter(XmlWriter w);

    /// <summary>
    /// A zip entry stamped with <see cref="Created"/>: without it the entry carries the clock time of the save and two
    /// saves of the same workbook differ in bytes. Zip times run from 1980 to 2107; outside that the clock is used.
    /// </summary>
    private ZipArchiveEntry Entry(ZipArchive zip, string name, CompressionLevel level)
    {
        var entry = zip.CreateEntry(name, level);
        if (Created is DateTime created)
        {
            var utc = created.ToUniversalTime();
            if (utc.Year >= 1980 && utc.Year <= 2107)
            {
                entry.LastWriteTime = new DateTimeOffset(DateTime.SpecifyKind(utc, DateTimeKind.Unspecified), TimeSpan.Zero);
            }
        }

        return entry;
    }

    private void Part(ZipArchive zip, string name, PartWriter write)
    {
        var entry = Entry(zip, name, CompressionLevel.Fastest);
        var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, CloseOutput = true };
        using (var w = XmlWriter.Create(entry.Open(), settings))
        {
            w.WriteStartDocument(true);
            write(w);
            w.WriteEndDocument();
        }
    }

    private void WriteSheet(XmlWriter w, XlsxSheet sheet, bool selected, StyleTable styles, StringTable strings)
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
        w.WriteAttributeString("fitToPage", sheet.FitWidth > 0 || sheet.FitHeight > 0 ? "1" : "0");
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

        // Rows with an explicit height are written even when they hold no cell.
        var heights = sheet.Heights.Keys.ToArray();
        Array.Sort(heights);

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
        w.WriteEndElement();

        // cols: only columns with a width or hidden.
        var columns = new SortedSet<int>(sheet.Widths.Keys);
        columns.UnionWith(sheet.HiddenColumns);
        if (columns.Count > 0)
        {
            w.WriteStartElement("cols", MainNs);
            foreach (var col in columns)
            {
                w.WriteStartElement("col", MainNs);
                w.WriteAttributeString("min", Int(col));
                w.WriteAttributeString("max", Int(col));
                w.WriteAttributeString("width", Num(sheet.Widths.TryGetValue(col, out var width) ? width : 8.43));
                if (sheet.HiddenColumns.Contains(col)) { w.WriteAttributeString("hidden", "1"); }
                w.WriteAttributeString("customWidth", "1");
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        // sheetData
        // Rows come from merging the sorted cell keys with the sorted heights. Addresses and numbers are formatted into
        // one scratch buffer: a sheet holds up to millions of cells, and a string for each would dominate the save.
        w.WriteStartElement("sheetData", MainNs);
        var buffer = new char[64];
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
                w.WriteAttributeString("ht", Num(sheet.Heights[row]));
                w.WriteAttributeString("customHeight", "1");
                nextHeight++;
            }

            while (index < count && XlsxSheet.RowOf(keys[index]) == row)
            {
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
                w.WriteAttributeString("password", PasswordHash(sheet.Password));
            }

            w.WriteAttributeString("sheet", "1");
            w.WriteAttributeString("objects", "1");
            w.WriteAttributeString("scenarios", "1");
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

        // Colour scales (conditional formatting): colours for the lowest / [median] / highest value.
        var priority = 0;
        foreach (var scale in sheet.ColorScales)
        {
            w.WriteStartElement("conditionalFormatting", MainNs);
            w.WriteAttributeString("sqref", CellName(scale.Row1, scale.Column1) + ":" + CellName(scale.Row2, scale.Column2));
            w.WriteStartElement("cfRule", MainNs);
            w.WriteAttributeString("type", "colorScale");
            w.WriteAttributeString("priority", Int(++priority));
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
            w.WriteEndElement();
            w.WriteEndElement();
        }

        // A link inside the workbook needs no relationship: location is enough.
        if (sheet.Links.Count > 0)
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

            w.WriteEndElement();
        }

        w.WriteStartElement("pageMargins", MainNs);
        w.WriteAttributeString("left", "0.5");
        w.WriteAttributeString("right", "0.5");
        w.WriteAttributeString("top", "0.5");
        w.WriteAttributeString("bottom", "0.5");
        w.WriteAttributeString("header", "0.3");
        w.WriteAttributeString("footer", "0.3");
        w.WriteEndElement();

        // paperSize 9 = A4.
        w.WriteStartElement("pageSetup", MainNs);
        w.WriteAttributeString("paperSize", "9");
        w.WriteAttributeString("orientation", sheet.Landscape ? "landscape" : "portrait");
        w.WriteAttributeString("fitToWidth", Int(sheet.FitWidth));
        w.WriteAttributeString("fitToHeight", Int(sheet.FitHeight));
        w.WriteEndElement();

        // Header and footer: sections &L &C &R, &P page, &N page count.
        if (sheet.Header != null || sheet.Footer != null)
        {
            w.WriteStartElement("headerFooter", MainNs);
            if (sheet.Header != null) { w.WriteElementString("oddHeader", MainNs, sheet.Header); }
            if (sheet.Footer != null) { w.WriteElementString("oddFooter", MainNs, sheet.Footer); }
            w.WriteEndElement();
        }

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

        // Charts and images: rId1 in the sheet relationships is the drawing layer.
        if (sheet.Charts.Count > 0 || sheet.Images.Count > 0)
        {
            w.WriteStartElement("drawing", MainNs);
            w.WriteAttributeString("id", RelNs, "rId1");
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private static void WriteCell(XmlWriter w, long key, XlsxSheet.Cell cell, StyleTable styles, StringTable strings, char[] buffer)
    {
        var style = styles.Index(cell.Style);
        if (cell.Value == null && style == 0)
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

        if (cell.Formula != null && cell.Value is string cached)
        {
            w.WriteAttributeString("t", "str");
            w.WriteElementString("f", MainNs, cell.Formula);
            w.WriteElementString("v", MainNs, cached);
        }
        else if (cell.Value is string text)
        {
            w.WriteAttributeString("t", "s");
            w.WriteStartElement("v", MainNs);
            w.WriteChars(buffer, 0, FormatInt(buffer, 0, strings.Index(text)));
            w.WriteEndElement();
        }
        else if (cell.Value is double number)
        {
            w.WriteStartElement("v", MainNs);
            w.WriteChars(buffer, 0, FormatDouble(buffer, number));
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private static void WriteNumberAttribute(XmlWriter w, string name, int value, char[] buffer)
    {
        w.WriteStartAttribute(name);
        w.WriteChars(buffer, 0, FormatInt(buffer, 0, value));
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

        return FormatInt(buffer, letters, row);
    }

    /// <summary>
    /// Writes a number into the buffer as <c>ToString("R")</c> would; returns its length.
    /// </summary>
    /// <remarks>
    /// Whole numbers below 10^15 (counters, years, IDs) are written digit by digit on every runtime: "R" gives them
    /// without exponent or point, so the text is the same. Other numbers go to the runtime's own round-trip format.
    /// </remarks>
    private static int FormatDouble(char[] buffer, double number)
    {
        // Negative zero is left to the runtime: .NET writes "-0", .NET Framework "0".
        if (number == Math.Floor(number) && Math.Abs(number) < 1e15 && (number != 0 || BitConverter.DoubleToInt64Bits(number) == 0))
        {
            var whole = (long)number;
            var position = 0;
            if (whole < 0)
            {
                buffer[position++] = '-';
                whole = -whole;
            }

            var digits = 1;
            for (var rest = whole; rest >= 10; rest /= 10)
            {
                digits++;
            }

            var end = position + digits;
            for (var i = end - 1; i >= position; i--)
            {
                buffer[i] = (char)('0' + (int)(whole % 10));
                whole /= 10;
            }

            return end;
        }

#if NET
        number.TryFormat(buffer, out var length, "R", CultureInfo.InvariantCulture);
        return length;
#else
        var text = number.ToString("R", CultureInfo.InvariantCulture);
        text.CopyTo(0, buffer, 0, text.Length);
        return text.Length;
#endif
    }

    /// <summary>
    /// Writes a non-negative whole number into the buffer from <paramref name="position"/>; returns the end position.
    /// </summary>
    private static int FormatInt(char[] buffer, int position, int value)
    {
        var digits = 1;
        for (var rest = value; rest >= 10; rest /= 10)
        {
            digits++;
        }

        var end = position + digits;
        for (var i = end - 1; i >= position; i--)
        {
            buffer[i] = (char)('0' + (value % 10));
            value /= 10;
        }

        return end;
    }

    private void WriteWorkbook(XmlWriter w)
    {
        w.WriteStartElement("workbook", MainNs);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        w.WriteStartElement("bookViews", MainNs);
        w.WriteStartElement("workbookView", MainNs);
        w.WriteAttributeString("activeTab", "0");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("sheets", MainNs);
        for (var i = 0; i < _sheets.Count; i++)
        {
            w.WriteStartElement("sheet", MainNs);
            w.WriteAttributeString("name", _sheets[i].Name);
            w.WriteAttributeString("sheetId", Int(i + 1));
            w.WriteAttributeString("id", RelNs, "rId" + (i + 1));
            w.WriteEndElement();
        }

        w.WriteEndElement();

        // Print area and print titles: reserved names bound to the sheet (_xlnm.Print_Area, _xlnm.Print_Titles).
        var named = _sheets.Where(s => s.PrintTitleRows != null || s.PrintArea != null).ToList();
        if (named.Count > 0)
        {
            w.WriteStartElement("definedNames", MainNs);
            foreach (var sheet in named)
            {
                var quoted = "'" + sheet.Name.Replace("'", "''") + "'!";
                if (sheet.PrintArea != null)
                {
                    var area = sheet.PrintArea.Value;
                    w.WriteStartElement("definedName", MainNs);
                    w.WriteAttributeString("name", "_xlnm.Print_Area");
                    w.WriteAttributeString("localSheetId", Int(_sheets.IndexOf(sheet)));
                    w.WriteString(quoted + "$" + ColumnName(area.FirstColumn) + "$" + Int(area.FirstRow) + ":$" + ColumnName(area.LastColumn) + "$" + Int(area.LastRow));
                    w.WriteEndElement();
                }

                if (sheet.PrintTitleRows != null)
                {
                    var rows = sheet.PrintTitleRows.Value;
                    w.WriteStartElement("definedName", MainNs);
                    w.WriteAttributeString("name", "_xlnm.Print_Titles");
                    w.WriteAttributeString("localSheetId", Int(_sheets.IndexOf(sheet)));
                    w.WriteString(quoted + "$" + Int(rows.First) + ":$" + Int(rows.Last));
                    w.WriteEndElement();
                }
            }

            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private void WriteWorkbookRels(XmlWriter w)
    {
        w.WriteStartElement("Relationships", PackageRelNs);
        for (var i = 0; i < _sheets.Count; i++)
        {
            Relationship(w, "rId" + (i + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet" + (i + 1) + ".xml");
        }

        Relationship(w, "rId" + (_sheets.Count + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
        Relationship(w, "rId" + (_sheets.Count + 2), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");
        w.WriteEndElement();
    }

    private static void WriteRootRels(XmlWriter w)
    {
        w.WriteStartElement("Relationships", PackageRelNs);
        Relationship(w, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
        Relationship(w, "rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
        Relationship(w, "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
        w.WriteEndElement();
    }

    private void WriteContentTypes(XmlWriter w)
    {
        w.WriteStartElement("Types", ContentTypesNs);
        Default(w, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        Default(w, "xml", "application/xml");
        Default(w, "png", "image/png");
        Override(w, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        for (var i = 0; i < _sheets.Count; i++)
        {
            Override(w, "/xl/worksheets/sheet" + (i + 1) + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        }

        for (var i = 1; i <= _drawings; i++)
        {
            Override(w, "/xl/drawings/drawing" + i + ".xml", "application/vnd.openxmlformats-officedocument.drawing+xml");
        }

        for (var i = 1; i <= _charts; i++)
        {
            Override(w, "/xl/charts/chart" + i + ".xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
        }

        Override(w, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        Override(w, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        Override(w, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
        Override(w, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
        w.WriteEndElement();
    }

    private void WriteCore(XmlWriter w)
    {
        const string Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        const string Dc = "http://purl.org/dc/elements/1.1/";
        const string Dcterms = "http://purl.org/dc/terms/";
        const string Xsi = "http://www.w3.org/2001/XMLSchema-instance";
        var now = (Created ?? DateTime.UtcNow).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        w.WriteStartElement("cp", "coreProperties", Cp);
        w.WriteAttributeString("xmlns", "dc", null, Dc);
        w.WriteAttributeString("xmlns", "dcterms", null, Dcterms);
        w.WriteAttributeString("xmlns", "xsi", null, Xsi);
        if (Title != null) { w.WriteElementString("dc", "title", Dc, Title); }
        if (Subject != null) { w.WriteElementString("dc", "subject", Dc, Subject); }
        w.WriteElementString("dc", "creator", Dc, Author ?? Application);
        if (Keywords != null) { w.WriteElementString("cp", "keywords", Cp, Keywords); }
        if (Description != null) { w.WriteElementString("dc", "description", Dc, Description); }
        foreach (var name in new[] { "created", "modified" })
        {
            w.WriteStartElement("dcterms", name, Dcterms);
            w.WriteAttributeString("xsi", "type", Xsi, "dcterms:W3CDTF");
            w.WriteString(now);
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private void WriteApp(XmlWriter w)
    {
        w.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        w.WriteElementString("Application", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", Application);
        w.WriteEndElement();
    }

    private static void Relationship(XmlWriter w, string id, string type, string target)
    {
        w.WriteStartElement("Relationship", PackageRelNs);
        w.WriteAttributeString("Id", id);
        w.WriteAttributeString("Type", type);
        w.WriteAttributeString("Target", target);
        w.WriteEndElement();
    }

    private static void Default(XmlWriter w, string extension, string type)
    {
        w.WriteStartElement("Default", ContentTypesNs);
        w.WriteAttributeString("Extension", extension);
        w.WriteAttributeString("ContentType", type);
        w.WriteEndElement();
    }

    private static void Override(XmlWriter w, string part, string type)
    {
        w.WriteStartElement("Override", ContentTypesNs);
        w.WriteAttributeString("PartName", part);
        w.WriteAttributeString("ContentType", type);
        w.WriteEndElement();
    }

    #endregion

    #region Helpers

    private static string Int(int value) => value.ToString(CultureInfo.InvariantCulture);

    private static void Cfvo(XmlWriter w, string type, string? value)
    {
        w.WriteStartElement("cfvo", MainNs);
        w.WriteAttributeString("type", type);
        if (value != null) { w.WriteAttributeString("val", value); }
        w.WriteEndElement();
    }

    /// <summary>
    /// Excel's legacy 16-bit password hash for sheet protection (ECMA-376, part 18.2.28); four hex digits.
    /// </summary>
    internal static string PasswordHash(string password)
    {
        var hash = 0;
        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= password[i];
        }

        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= password.Length;
        hash ^= 0xCE4B;
        return hash.ToString("X4", CultureInfo.InvariantCulture);
    }

    private static string Num(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    /// <summary>
    /// <c>#RRGGBB</c> → Excel's <c>FFRRGGBB</c>.
    /// </summary>
    internal static string Argb(string color)
    {
        var hex = color.TrimStart('#').ToUpperInvariant();
        if (hex.Length != 6)
        {
            throw new ArgumentException(Messages.Format(Strings.ColorInvalid, color));
        }

        return "FF" + hex;
    }

    #endregion

    #region Tables

    /// <summary>
    /// Shared strings: each distinct text once, in order of first use.
    /// </summary>
    private sealed class StringTable
    {
        private readonly Dictionary<string, int> _index = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<string> _texts = new List<string>();
        private int _uses;

        public int Index(string text)
        {
            _uses++;
            if (!_index.TryGetValue(text, out var i))
            {
                i = _texts.Count;
                _texts.Add(text);
                _index[text] = i;
            }

            return i;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("sst", MainNs);
            w.WriteAttributeString("count", Int(_uses));
            w.WriteAttributeString("uniqueCount", Int(_texts.Count));
            foreach (var text in _texts)
            {
                w.WriteStartElement("si", MainNs);
                w.WriteStartElement("t", MainNs);
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]) || text.IndexOf('\n') >= 0))
                {
                    w.WriteAttributeString("xml", "space", null, "preserve");
                }

                w.WriteString(Clean(text));
                w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        /// <summary>
        /// Drops control characters that XML cannot hold (tab and line feed stay); XmlWriter throws on them.
        /// </summary>
        private static string Clean(string text)
        {
            StringBuilder? sb = null;
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                var bad = ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r';
                if (bad && sb == null)
                {
                    sb = new StringBuilder(text, 0, i, text.Length);
                }
                else if (!bad)
                {
                    sb?.Append(ch);
                }
            }

            return sb?.ToString() ?? text;
        }
    }

    /// <summary>
    /// Style table: each distinct <see cref="XlsxStyle"/> is one <c>xf</c> record; fonts, fills, borders and number
    /// formats are made unique separately. Record 0 is the default style.
    /// </summary>
    private sealed class StyleTable
    {
        private readonly XlsxWorkbook _book;
        private readonly Dictionary<XlsxStyle, int> _xfIndex = new Dictionary<XlsxStyle, int>();
        private readonly List<(int Font, int Fill, int Border, int NumFmt, XlsxStyle Style)> _xfs = new List<(int, int, int, int, XlsxStyle)>();
        private readonly List<string> _fonts = new List<string>();
        private readonly List<string> _fills = new List<string>();
        private readonly List<string> _borders = new List<string>();
        private readonly List<string> _numFmts = new List<string>();

        /// <summary>
        /// The style asked for last and its record; starts as the default style, record 0.
        /// </summary>
        private XlsxStyle _last;

        private int _lastIndex;

        public StyleTable(XlsxWorkbook book)
        {
            _book = book;

            // Required start of the standard: font 0 is the default, fills 0 and 1 are none and gray125, border 0 is empty.
            _fonts.Add(FontKey(default(XlsxStyle)));
            _fills.Add("none");
            _fills.Add("gray125");
            _borders.Add(BorderKey(default(XlsxStyle)));
            _xfIndex[default(XlsxStyle)] = 0;
            _xfs.Add((0, 0, 0, 0, default(XlsxStyle)));
        }

        public int Index(XlsxStyle style)
        {
            // Neighbouring cells usually share a style; comparing with the last one is cheaper than hashing every string.
            if (style.Equals(_last))
            {
                return _lastIndex;
            }

            _last = style;
            if (_xfIndex.TryGetValue(style, out var i))
            {
                _lastIndex = i;
                return i;
            }

            var font = Add(_fonts, FontKey(style));
            var fill = style.Fill == null ? 0 : Add(_fills, "solid:" + Argb(style.Fill));
            var border = Add(_borders, BorderKey(style));
            var numFmt = style.NumberFormat == null ? 0 : 164 + Add(_numFmts, style.NumberFormat);
            i = _xfs.Count;
            _xfs.Add((font, fill, border, numFmt, style));
            _xfIndex[style] = i;
            _lastIndex = i;
            return i;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("styleSheet", MainNs);

            if (_numFmts.Count > 0)
            {
                w.WriteStartElement("numFmts", MainNs);
                w.WriteAttributeString("count", Int(_numFmts.Count));
                for (var i = 0; i < _numFmts.Count; i++)
                {
                    w.WriteStartElement("numFmt", MainNs);
                    w.WriteAttributeString("numFmtId", Int(164 + i));
                    w.WriteAttributeString("formatCode", _numFmts[i]);
                    w.WriteEndElement();
                }

                w.WriteEndElement();
            }

            // Fonts: key "name|size|B|I|U|colour".
            w.WriteStartElement("fonts", MainNs);
            w.WriteAttributeString("count", Int(_fonts.Count));
            foreach (var key in _fonts)
            {
                var p = key.Split('|');
                w.WriteStartElement("font", MainNs);
                if (p[2] == "1") { w.WriteElementString("b", MainNs, null); }
                if (p[3] == "1") { w.WriteElementString("i", MainNs, null); }
                if (p[4] == "1") { w.WriteElementString("u", MainNs, null); }
                w.WriteStartElement("sz", MainNs); w.WriteAttributeString("val", p[1]); w.WriteEndElement();
                if (p[5].Length > 0) { w.WriteStartElement("color", MainNs); w.WriteAttributeString("rgb", p[5]); w.WriteEndElement(); }
                w.WriteStartElement("name", MainNs); w.WriteAttributeString("val", p[0]); w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();

            w.WriteStartElement("fills", MainNs);
            w.WriteAttributeString("count", Int(_fills.Count));
            foreach (var key in _fills)
            {
                w.WriteStartElement("fill", MainNs);
                w.WriteStartElement("patternFill", MainNs);
                if (key.StartsWith("solid:", StringComparison.Ordinal))
                {
                    w.WriteAttributeString("patternType", "solid");
                    w.WriteStartElement("fgColor", MainNs); w.WriteAttributeString("rgb", key.Substring(6)); w.WriteEndElement();
                    w.WriteStartElement("bgColor", MainNs); w.WriteAttributeString("indexed", "64"); w.WriteEndElement();
                }
                else
                {
                    w.WriteAttributeString("patternType", key);
                }

                w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();

            // Borders: key "left|right|top|bottom", each "kind,colour".
            w.WriteStartElement("borders", MainNs);
            w.WriteAttributeString("count", Int(_borders.Count));
            foreach (var key in _borders)
            {
                var p = key.Split('|');
                w.WriteStartElement("border", MainNs);
                Side(w, "left", p[0]);
                Side(w, "right", p[1]);
                Side(w, "top", p[2]);
                Side(w, "bottom", p[3]);
                w.WriteStartElement("diagonal", MainNs); w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();

            w.WriteStartElement("cellStyleXfs", MainNs);
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("xf", MainNs);
            w.WriteAttributeString("numFmtId", "0"); w.WriteAttributeString("fontId", "0"); w.WriteAttributeString("fillId", "0"); w.WriteAttributeString("borderId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("cellXfs", MainNs);
            w.WriteAttributeString("count", Int(_xfs.Count));
            foreach (var xf in _xfs)
            {
                var s = xf.Style;
                w.WriteStartElement("xf", MainNs);
                w.WriteAttributeString("numFmtId", Int(xf.NumFmt));
                w.WriteAttributeString("fontId", Int(xf.Font));
                w.WriteAttributeString("fillId", Int(xf.Fill));
                w.WriteAttributeString("borderId", Int(xf.Border));
                w.WriteAttributeString("xfId", "0");
                if (xf.NumFmt != 0) { w.WriteAttributeString("applyNumberFormat", "1"); }
                if (xf.Font != 0) { w.WriteAttributeString("applyFont", "1"); }
                if (xf.Fill != 0) { w.WriteAttributeString("applyFill", "1"); }
                if (xf.Border != 0) { w.WriteAttributeString("applyBorder", "1"); }
                var aligned = s.Horizontal != null || s.Vertical != null || s.Wrap;
                if (aligned)
                {
                    w.WriteAttributeString("applyAlignment", "1");
                    w.WriteStartElement("alignment", MainNs);
                    if (s.Horizontal != null) { w.WriteAttributeString("horizontal", s.Horizontal); }
                    if (s.Vertical != null) { w.WriteAttributeString("vertical", s.Vertical); }
                    if (s.Wrap) { w.WriteAttributeString("wrapText", "1"); }
                    w.WriteEndElement();
                }

                w.WriteEndElement();
            }

            w.WriteEndElement();

            w.WriteStartElement("cellStyles", MainNs);
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("cellStyle", MainNs);
            w.WriteAttributeString("name", "Normal"); w.WriteAttributeString("xfId", "0"); w.WriteAttributeString("builtinId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteEndElement();
        }

        private string FontKey(XlsxStyle s) =>
            (s.FontName ?? _book.DefaultFontName) + "|" + Num(s.FontSize > 0 ? s.FontSize : _book.DefaultFontSize) + "|"
            + (s.Bold ? "1" : "0") + "|" + (s.Italic ? "1" : "0") + "|" + (s.Underline ? "1" : "0") + "|"
            + (s.FontColor == null ? string.Empty : Argb(s.FontColor));

        private static string BorderKey(XlsxStyle s) => Edge(s.Left) + "|" + Edge(s.Right) + "|" + Edge(s.Top) + "|" + Edge(s.Bottom);

        private static string Edge(XlsxBorder b) => b.Line == null ? string.Empty : b.Line + "," + (b.Color == null ? string.Empty : Argb(b.Color));

        private static void Side(XmlWriter w, string name, string key)
        {
            w.WriteStartElement(name, MainNs);
            if (key.Length > 0)
            {
                var p = key.Split(',');
                w.WriteAttributeString("style", p[0]);
                if (p[1].Length > 0)
                {
                    w.WriteStartElement("color", MainNs);
                    w.WriteAttributeString("rgb", p[1]);
                    w.WriteEndElement();
                }
            }

            w.WriteEndElement();
        }

        private static int Add(List<string> list, string key)
        {
            var i = list.IndexOf(key);
            if (i < 0)
            {
                i = list.Count;
                list.Add(key);
            }

            return i;
        }
    }

    #endregion
}
