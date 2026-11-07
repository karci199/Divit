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
    /// Defined names in the order they were added.
    /// </summary>
    private readonly List<(string Name, XlsxSheet Sheet, int Row1, int Col1, int Row2, int Col2, XlsxSheet? Scope)> _names =
        new List<(string, XlsxSheet, int, int, int, int, XlsxSheet?)>();

    /// <summary>
    /// Number of drawing layers and charts written by the last save; the content type list is built from them.
    /// </summary>
    private int _drawings;

    private int _charts;

    /// <summary>
    /// Number of sheets with comments in the last save, and the tab index of the sheet that opens first.
    /// </summary>
    private int _notes;

    private int _active;

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

    /// <summary>
    /// The sheet that is open when the file opens; the first visible sheet when empty. It must be visible.
    /// </summary>
    public XlsxSheet? ActiveSheet { get; set; }

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
    /// Moves a sheet to another place in the tab order.
    /// </summary>
    /// <param name="sheet">A sheet of this workbook.</param>
    /// <param name="index">New position, 0 for the first tab.</param>
    /// <exception cref="ArgumentException">The sheet belongs to another workbook.</exception>
    public void MoveSheet(XlsxSheet sheet, int index)
    {
        if (!_sheets.Remove(sheet))
        {
            throw new ArgumentException(Messages.Format(Strings.SheetNotInWorkbook, sheet.Name), nameof(sheet));
        }

        _sheets.Insert(Math.Max(0, Math.Min(index, _sheets.Count)), sheet);
    }

    /// <summary>
    /// Gives a range a name, usable in formulas (<c>=SUM(Sales)</c>) and in Excel's Name Box to jump to it.
    /// </summary>
    /// <param name="name">The name: starts with a letter or underscore, then letters, digits, underscores and dots;
    /// not a cell address such as <c>A1</c> or <c>R1C1</c>; unique in its scope (ignoring case).</param>
    /// <param name="sheet">Sheet of the range.</param>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="scope">The sheet the name belongs to, so the same name can exist on other sheets; the whole
    /// workbook when empty.</param>
    /// <exception cref="ArgumentException">The name is not usable or the sheet belongs to another workbook.</exception>
    public void DefineName(string name, XlsxSheet sheet, int row1, int column1, int row2, int column2, XlsxSheet? scope = null)
    {
        foreach (var owned in new[] { sheet, scope })
        {
            if (owned != null && !_sheets.Contains(owned))
            {
                throw new ArgumentException(Messages.Format(Strings.SheetNotInWorkbook, owned.Name), nameof(sheet));
            }
        }

        if (!IsDefinedName(name))
        {
            throw new ArgumentException(Messages.Format(Strings.DefinedNameInvalid, name), nameof(name));
        }

        foreach (var defined in _names)
        {
            if (defined.Scope == scope && string.Equals(defined.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(Messages.Format(Strings.DefinedNameDuplicate, name), nameof(name));
            }
        }

        if (row1 < 1 || column1 < 1 || row2 > XlsxSheet.MaxRow || column2 > XlsxSheet.MaxColumn || row2 < row1 || column2 < column1)
        {
            throw new ArgumentOutOfRangeException(nameof(row1), Messages.Format(Strings.CellOutOfRange, row1, column1));
        }

        _names.Add((name, sheet, row1, column1, row2, column2, scope));
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

        var active = ActiveSheet ?? _sheets.FirstOrDefault(s => s.Visibility == XlsxSheetVisibility.Visible);
        if (active == null)
        {
            throw new InvalidOperationException(Messages.Get(Strings.NoVisibleSheet));
        }

        if (!_sheets.Contains(active))
        {
            throw new InvalidOperationException(Messages.Format(Strings.SheetNotInWorkbook, active.Name));
        }

        if (active.Visibility != XlsxSheetVisibility.Visible)
        {
            throw new InvalidOperationException(Messages.Format(Strings.ActiveSheetHidden, active.Name));
        }

        _active = _sheets.IndexOf(active);
        var plans = PlanSheets();
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
                    Part(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", w => WriteSheet(w, _sheets[i], i == _active, plans[i], styles, strings));
                }

                // The sheets' relationships, then their notes: a comment list and the legacy VML shapes Excel draws them with.
                for (var i = 0; i < _sheets.Count; i++)
                {
                    var plan = plans[i];
                    if (plan.Relationships.Count > 0)
                    {
                        Part(zip, "xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels", w =>
                        {
                            w.WriteStartElement("Relationships", PackageRelNs);
                            foreach (var rel in plan.Relationships)
                            {
                                Relationship(w, rel.Id, rel.Type, rel.Target, rel.External);
                            }

                            w.WriteEndElement();
                        });
                    }

                    if (plan.Comments > 0)
                    {
                        var sheet = _sheets[i];
                        Part(zip, "xl/comments" + plan.Comments + ".xml", w => WriteComments(w, sheet));
                        Part(zip, "xl/drawings/vmlDrawing" + plan.Comments + ".vml", w => WriteNoteShapes(w, sheet, plan.Comments), false);
                    }
                }

                // Every sheet with charts or images: sheet → drawing layer → charts and images. Numbers run across the workbook.
                _charts = 0;
                var media = 0;
                for (var i = 0; i < _sheets.Count; i++)
                {
                    var charts = _sheets[i].Charts;
                    var images = _sheets[i].Images;
                    if (plans[i].Drawing == 0)
                    {
                        continue;
                    }

                    var drawing = plans[i].Drawing;
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

    /// <summary>
    /// Writes one XML part; <paramref name="declaration"/> is off only for the legacy VML part, which Excel writes
    /// without one.
    /// </summary>
    private void Part(ZipArchive zip, string name, PartWriter write, bool declaration = true)
    {
        var entry = Entry(zip, name, CompressionLevel.Fastest);
        var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, CloseOutput = true, OmitXmlDeclaration = !declaration };
        using (var w = XmlWriter.Create(entry.Open(), settings))
        {
            if (declaration)
            {
                w.WriteStartDocument(true);
            }

            write(w);
            w.WriteEndDocument();
        }
    }

    /// <summary>
    /// Numbers the parts of every sheet and gives their relationships ids, before anything is written: the sheet XML
    /// names the ids, the relationship parts and content types list the parts.
    /// </summary>
    private SheetPlan[] PlanSheets()
    {
        const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        var plans = new SheetPlan[_sheets.Count];
        _drawings = 0;
        _notes = 0;
        for (var i = 0; i < _sheets.Count; i++)
        {
            var sheet = _sheets[i];
            var plan = new SheetPlan();
            if (sheet.Charts.Count > 0 || sheet.Images.Count > 0)
            {
                plan.Drawing = ++_drawings;
                plan.DrawingId = plan.Add(Rel + "drawing", "../drawings/drawing" + plan.Drawing + ".xml", false);
            }

            if (sheet.Comments.Count > 0)
            {
                plan.Comments = ++_notes;
                plan.Add(Rel + "comments", "../comments" + plan.Comments + ".xml", false);
                plan.NoteShapesId = plan.Add(Rel + "vmlDrawing", "../drawings/vmlDrawing" + plan.Comments + ".vml", false);
            }

            foreach (var link in sheet.UrlLinks)
            {
                plan.LinkIds.Add(plan.Add(Rel + "hyperlink", link.Url, true));
            }

            plans[i] = plan;
        }

        return plans;
    }

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

        // Rows with a height, hidden or grouped are written even when they hold no cell.
        var heights = sheet.Rows.Keys.ToArray();
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
            w.WriteElementString("author", MainNs, StringTable.Clean(author));
        }

        w.WriteEndElement();
        w.WriteStartElement("commentList", MainNs);
        foreach (var comment in sheet.Comments)
        {
            w.WriteStartElement("comment", MainNs);
            w.WriteAttributeString("ref", CellName(comment.Row, comment.Column));
            w.WriteAttributeString("authorId", Int(authors.IndexOf(comment.Author ?? Author ?? Application)));
            w.WriteStartElement("text", MainNs);
            w.WriteStartElement("t", MainNs);
            w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(StringTable.Clean(comment.Text));
            w.WriteEndElement();
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
                Int(column + 1) + ", 15, " + Int(Math.Max(0, row - 1)) + ", 10, " + Int(column + 3) + ", 15, " + Int(Math.Max(0, row - 1) + 4) + ", 4");
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
        else if (cell.Value is bool flag)
        {
            w.WriteAttributeString("t", "b");
            w.WriteElementString("v", MainNs, flag ? "1" : "0");
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
        // firstSheet is the leftmost tab shown in the tab bar; a hidden one there leaves the bar looking empty.
        w.WriteStartElement("workbookView", MainNs);
        var first = _sheets.FindIndex(s => s.Visibility == XlsxSheetVisibility.Visible);
        if (first > 0) { w.WriteAttributeString("firstSheet", Int(first)); }
        w.WriteAttributeString("activeTab", Int(_active));
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("sheets", MainNs);
        for (var i = 0; i < _sheets.Count; i++)
        {
            w.WriteStartElement("sheet", MainNs);
            w.WriteAttributeString("name", _sheets[i].Name);
            w.WriteAttributeString("sheetId", Int(i + 1));
            if (_sheets[i].Visibility == XlsxSheetVisibility.Hidden) { w.WriteAttributeString("state", "hidden"); }
            if (_sheets[i].Visibility == XlsxSheetVisibility.VeryHidden) { w.WriteAttributeString("state", "veryHidden"); }
            w.WriteAttributeString("id", RelNs, "rId" + (i + 1));
            w.WriteEndElement();
        }

        w.WriteEndElement();

        // The caller's names, then print area and print titles: reserved names bound to the sheet (_xlnm.Print_Area,
        // _xlnm.Print_Titles).
        var named = _sheets.Where(s => s.PrintTitleRows != null || s.PrintArea != null).ToList();
        if (named.Count > 0 || _names.Count > 0)
        {
            w.WriteStartElement("definedNames", MainNs);
            foreach (var name in _names)
            {
                w.WriteStartElement("definedName", MainNs);
                w.WriteAttributeString("name", name.Name);
                if (name.Scope != null) { w.WriteAttributeString("localSheetId", Int(_sheets.IndexOf(name.Scope))); }
                w.WriteString("'" + name.Sheet.Name.Replace("'", "''") + "'!$" + ColumnName(name.Col1) + "$" + Int(name.Row1)
                    + (name.Row1 == name.Row2 && name.Col1 == name.Col2 ? string.Empty : ":$" + ColumnName(name.Col2) + "$" + Int(name.Row2)));
                w.WriteEndElement();
            }

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
        if (_notes > 0)
        {
            Default(w, "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing");
        }

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

        for (var i = 1; i <= _notes; i++)
        {
            Override(w, "/xl/comments" + i + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
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

    private static void Relationship(XmlWriter w, string id, string type, string target, bool external = false)
    {
        w.WriteStartElement("Relationship", PackageRelNs);
        w.WriteAttributeString("Id", id);
        w.WriteAttributeString("Type", type);
        w.WriteAttributeString("Target", target);
        if (external) { w.WriteAttributeString("TargetMode", "External"); }
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

    /// <summary>
    /// Excel's rules for a defined name: letter, underscore or backslash first, then letters, digits, underscores, dots
    /// and backslashes; at most 255 characters; not a cell address in either notation; not a reserved name.
    /// </summary>
    private static bool IsDefinedName(string name)
    {
        if (name.Length == 0 || name.Length > 255 || name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!(char.IsLetter(name[0]) || name[0] == '_' || name[0] == '\\'))
        {
            return false;
        }

        foreach (var ch in name)
        {
            if (!(char.IsLetterOrDigit(ch) || ch == '_' || ch == '.' || ch == '\\'))
            {
                return false;
            }
        }

        return !LooksLikeA1(name) && !LooksLikeR1C1(name);
    }

    /// <summary>
    /// Letters of an existing column (up to XFD) followed by digits: <c>B3</c>, <c>XFD1</c>.
    /// </summary>
    private static bool LooksLikeA1(string name)
    {
        var letters = 0;
        var column = 0;
        while (letters < name.Length && letters < 4 && IsAsciiLetter(name[letters]))
        {
            column = (column * 26) + (char.ToUpperInvariant(name[letters]) - 'A' + 1);
            letters++;
        }

        if (letters == 0 || letters == name.Length || column > XlsxSheet.MaxColumn)
        {
            return false;
        }

        for (var i = letters; i < name.Length; i++)
        {
            if (name[i] < '0' || name[i] > '9')
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// <c>R</c>, <c>C</c>, <c>R1</c>, <c>C2</c>, <c>R1C2</c> and the like, in any case.
    /// </summary>
    private static bool LooksLikeR1C1(string name)
    {
        var i = 0;
        var parts = 0;
        foreach (var letter in new[] { 'R', 'C' })
        {
            if (i < name.Length && char.ToUpperInvariant(name[i]) == letter)
            {
                i++;
                parts++;
                while (i < name.Length && name[i] >= '0' && name[i] <= '9')
                {
                    i++;
                }
            }
        }

        return parts > 0 && i == name.Length;
    }

    private static bool IsAsciiLetter(char ch) => (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');

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
    /// A margin in inches; four decimals keep margins given in centimetres exact to a hundredth of a millimetre.
    /// </summary>
    private static string Inches(double value) => Math.Max(0, value).ToString("0.####", CultureInfo.InvariantCulture);

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
    /// Parts and relationship ids of one sheet; numbers 0 mean the sheet has no such part.
    /// </summary>
    private sealed class SheetPlan
    {
        public readonly List<(string Id, string Type, string Target, bool External)> Relationships =
            new List<(string, string, string, bool)>();

        /// <summary>
        /// Relationship ids of the URL links, in the order of <see cref="XlsxSheet.UrlLinks"/>.
        /// </summary>
        public readonly List<string> LinkIds = new List<string>();

        public int Drawing;

        public int Comments;

        public string? DrawingId;

        public string? NoteShapesId;

        public string Add(string type, string target, bool external)
        {
            var id = "rId" + Int(Relationships.Count + 1);
            Relationships.Add((id, type, target, external));
            return id;
        }
    }

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
        public static string Clean(string text)
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

            // Fonts: key "name|size|bold|italic|underline|colour|strike|script"; underline 1 single, 2 double; script 1
            // superscript, 2 subscript. Children in the order Excel writes them.
            w.WriteStartElement("fonts", MainNs);
            w.WriteAttributeString("count", Int(_fonts.Count));
            foreach (var key in _fonts)
            {
                var p = key.Split('|');
                w.WriteStartElement("font", MainNs);
                if (p[2] == "1") { w.WriteElementString("b", MainNs, null); }
                if (p[3] == "1") { w.WriteElementString("i", MainNs, null); }
                if (p[6] == "1") { w.WriteElementString("strike", MainNs, null); }
                if (p[4] == "1") { w.WriteElementString("u", MainNs, null); }
                if (p[4] == "2") { w.WriteStartElement("u", MainNs); w.WriteAttributeString("val", "double"); w.WriteEndElement(); }
                if (p[7] != "0")
                {
                    w.WriteStartElement("vertAlign", MainNs);
                    w.WriteAttributeString("val", p[7] == "1" ? "superscript" : "subscript");
                    w.WriteEndElement();
                }

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
                var aligned = s.Horizontal != null || s.Vertical != null || s.Wrap || s.ShrinkToFit || s.Indent > 0 || s.Rotation != 0;
                if (aligned)
                {
                    w.WriteAttributeString("applyAlignment", "1");
                    w.WriteStartElement("alignment", MainNs);

                    // An indent needs an edge to count from; Excel itself picks left.
                    var horizontal = s.Horizontal ?? (s.Indent > 0 ? "left" : null);
                    if (horizontal != null) { w.WriteAttributeString("horizontal", horizontal); }
                    if (s.Vertical != null) { w.WriteAttributeString("vertical", s.Vertical); }
                    if (s.Rotation != 0) { w.WriteAttributeString("textRotation", Int(TextRotation(s.Rotation))); }
                    if (s.Wrap) { w.WriteAttributeString("wrapText", "1"); }
                    if (s.Indent > 0) { w.WriteAttributeString("indent", Int(Math.Min(s.Indent, 250))); }
                    if (s.ShrinkToFit) { w.WriteAttributeString("shrinkToFit", "1"); }
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
            + (s.Bold ? "1" : "0") + "|" + (s.Italic ? "1" : "0") + "|" + (s.DoubleUnderline ? "2" : s.Underline ? "1" : "0") + "|"
            + (s.FontColor == null ? string.Empty : Argb(s.FontColor)) + "|" + (s.Strikethrough ? "1" : "0") + "|"
            + Int((int)s.Script);

        /// <summary>
        /// Degrees to Excel's textRotation: 0–90 counter-clockwise, 91–180 for 1–90 clockwise, 255 stacked.
        /// </summary>
        private static int TextRotation(int degrees)
        {
            if (degrees == XlsxStyle.VerticalText)
            {
                return degrees;
            }

            degrees = Math.Max(-90, Math.Min(90, degrees));
            return degrees >= 0 ? degrees : 90 - degrees;
        }

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
