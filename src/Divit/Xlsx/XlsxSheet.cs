#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;
using System.Collections.Generic;

/// <summary>
/// One sheet of a workbook: cells, merges, sizes, links, view and print settings.
/// </summary>
/// <remarks>
/// Row and column numbers start at 1 (A1 is 1, 1). Cells are stored sparsely: only a cell with a value or a style
/// takes space. The sheet only collects; turning it into XML happens in <see cref="XlsxWorkbook.Save"/>.
/// </remarks>
public sealed class XlsxSheet
{
    #region Fields

    /// <summary>
    /// Excel's last column (XFD).
    /// </summary>
    internal const int MaxColumn = 16384;

    /// <summary>
    /// Excel's last row.
    /// </summary>
    internal const int MaxRow = 1048576;

    private readonly Dictionary<long, Cell> _cells = new Dictionary<long, Cell>();
    private readonly Dictionary<int, double> _widths = new Dictionary<int, double>();
    private readonly Dictionary<int, double> _heights = new Dictionary<int, double>();
    private readonly HashSet<int> _hidden = new HashSet<int>();
    private readonly List<(int Row1, int Col1, int Row2, int Col2)> _merges = new List<(int, int, int, int)>();
    private readonly List<(int Row, int Col, string Sheet)> _links = new List<(int, int, string)>();
    private readonly List<XlsxChart> _charts = new List<XlsxChart>();
    private readonly List<XlsxImage> _images = new List<XlsxImage>();
    private readonly List<XlsxColorScale> _colorScales = new List<XlsxColorScale>();

    #endregion

    #region Construction

    internal XlsxSheet(string name)
    {
        Name = name;
    }

    #endregion

    #region Properties

    /// <summary>
    /// Name on the sheet tab.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Tab colour as <c>#RRGGBB</c>; no colour when empty.
    /// </summary>
    public string? TabColor { get; set; }

    /// <summary>
    /// Whether Excel's grey gridlines are shown.
    /// </summary>
    public bool ShowGridLines { get; set; } = true;

    /// <summary>
    /// Number of frozen rows at the top; 0 for none.
    /// </summary>
    public int FrozenRows { get; set; }

    /// <summary>
    /// Number of frozen columns on the left; 0 for none.
    /// </summary>
    public int FrozenColumns { get; set; }

    /// <summary>
    /// Prints in landscape orientation.
    /// </summary>
    public bool Landscape { get; set; }

    /// <summary>
    /// Number of pages the width is fitted to when printing; 0 turns fitting off.
    /// </summary>
    public int FitWidth { get; set; } = 1;

    /// <summary>
    /// Number of pages the height is fitted to when printing; 0 uses as many pages as needed.
    /// </summary>
    public int FitHeight { get; set; }

    /// <summary>
    /// Rows repeated at the top of every printed page (for example a table header); <c>null</c> for none.
    /// </summary>
    public (int First, int Last)? PrintTitleRows { get; set; }

    /// <summary>
    /// The rectangle to print (rows and columns from 1, ends included); <c>null</c> prints the whole sheet.
    /// </summary>
    public (int FirstRow, int FirstColumn, int LastRow, int LastColumn)? PrintArea { get; set; }

    /// <summary>
    /// Page header with Excel codes: <c>&amp;L</c> left, <c>&amp;C</c> centre, <c>&amp;R</c> right, <c>&amp;P</c> page
    /// number, <c>&amp;N</c> page count, <c>&amp;A</c> tab name. None when empty.
    /// </summary>
    public string? Header { get; set; }

    /// <summary>
    /// Page footer with the same codes as <see cref="Header"/>, for example <c>&amp;LDivit&amp;RPage &amp;P / &amp;N</c>.
    /// </summary>
    public string? Footer { get; set; }

    /// <summary>
    /// Whether to hide Excel's green "number stored as text" warning on text that looks like a number. Meant for fields
    /// written as text on purpose (serial numbers, where a leading zero must stay); applies to the whole sheet.
    /// </summary>
    public bool IgnoreNumbersAsText { get; set; }

    /// <summary>
    /// View zoom in Excel, percent (10–400); 0 uses <see cref="XlsxWorkbook.DefaultZoom"/>.
    /// </summary>
    public int Zoom { get; set; }

    /// <summary>
    /// Whether the sheet is locked: cells cannot be changed, links still work. See <see cref="Password"/> for unlocking.
    /// </summary>
    public bool Protected { get; private set; }

    /// <summary>
    /// Password of the lock; when empty the lock has no password (Excel removes it with one click; it only guards
    /// against accidental edits).
    /// </summary>
    public string? Password { get; private set; }

    /// <summary>
    /// The sheet's charts in the order they were added.
    /// </summary>
    public IReadOnlyList<XlsxChart> Charts => _charts;

    internal IReadOnlyList<XlsxImage> Images => _images;

    internal IReadOnlyList<XlsxColorScale> ColorScales => _colorScales;

    internal IReadOnlyCollection<int> HiddenColumns => _hidden;

    internal (int Row1, int Col1, int Row2, int Col2)? AutoFilterRange { get; private set; }

    internal IEnumerable<KeyValuePair<long, Cell>> Cells => _cells;

    internal IReadOnlyDictionary<int, double> Widths => _widths;

    internal IReadOnlyDictionary<int, double> Heights => _heights;

    internal IReadOnlyList<(int Row1, int Col1, int Row2, int Col2)> Merges => _merges;

    internal IReadOnlyList<(int Row, int Col, string Sheet)> Links => _links;

    #endregion

    #region Protection

    /// <summary>
    /// Locks the sheet: values cannot be edited by hand; selecting, copying and clicking links still work.
    /// </summary>
    /// <remarks>
    /// Uses Excel's legacy 16-bit password hash. It is not authentication, only protection against accidental changes;
    /// to prove that a file was not modified, keep a hash of the whole file elsewhere.
    /// </remarks>
    /// <param name="password">Password; none when empty.</param>
    public void Protect(string? password = null)
    {
        Protected = true;
        Password = password;
    }

    #endregion

    #region Images and colour scales

    /// <summary>
    /// Adds a PNG image to the sheet (for example a heat map or a chart drawn by another program).
    /// </summary>
    /// <param name="row">Row of the top-left corner.</param>
    /// <param name="column">Column of the top-left corner.</param>
    /// <param name="png">Bytes of the PNG file.</param>
    /// <param name="width">Width in pixels.</param>
    /// <param name="height">Height in pixels.</param>
    /// <exception cref="ArgumentException">The bytes are not a PNG image.</exception>
    public void AddImage(int row, int column, byte[] png, int width, int height)
    {
        Check(row, column);
        if (png.Length < 8 || png[0] != 0x89 || png[1] != 0x50 || png[2] != 0x4E || png[3] != 0x47)
        {
            throw new ArgumentException(Messages.Get(Strings.ImageNotPng), nameof(png));
        }

        _images.Add(new XlsxImage(row, column, png, width, height));
    }

    /// <summary>
    /// Gives a range a colour scale: each cell is coloured by its value (Excel conditional formatting). On 2D data this
    /// is an in-cell heat map: the values stay readable and the colour shows their size.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="lowColor">Colour of the lowest value.</param>
    /// <param name="midColor">Colour of the median; a two-colour scale when empty.</param>
    /// <param name="highColor">Colour of the highest value.</param>
    public void ColorScale(int row1, int column1, int row2, int column2, string lowColor, string? midColor, string highColor)
    {
        Check(row1, column1);
        Check(row2, column2);
        _colorScales.Add(new XlsxColorScale(row1, column1, row2, column2, lowColor, midColor, highColor));
    }

    #endregion

    #region Cells

    /// <summary>
    /// Writes text into a cell.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="text">The text.</param>
    public void Set(int row, int column, string text) => At(row, column).Value = text;

    /// <summary>
    /// Writes a number into a cell. NaN and infinity are not numbers in Excel; the cell stays empty.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="number">The number.</param>
    public void Set(int row, int column, double number) =>
        At(row, column).Value = double.IsNaN(number) || double.IsInfinity(number) ? null : (object)number;

    /// <summary>
    /// Changes the style of every cell in a range; the edit stacks on top of earlier ones.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="edit">The change to apply.</param>
    public void Format(int row1, int column1, int row2, int column2, XlsxStyleEdit edit)
    {
        for (var r = row1; r <= row2; r++)
        {
            for (var c = column1; c <= column2; c++)
            {
                edit(ref At(r, c).Style);
            }
        }
    }

    /// <summary>
    /// Changes the style of one cell.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="edit">The change to apply.</param>
    public void Format(int row, int column, XlsxStyleEdit edit) => Format(row, column, row, column, edit);

    /// <summary>
    /// Gives a range one style, replacing the previous one; for callers that build the style elsewhere.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="style">The style.</param>
    public void SetStyle(int row1, int column1, int row2, int column2, XlsxStyle style)
    {
        for (var r = row1; r <= row2; r++)
        {
            for (var c = column1; c <= column2; c++)
            {
                At(r, c).Style = style;
            }
        }
    }

    /// <summary>
    /// Draws the outer frame of a range and, if asked, its inner lines.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="outer">Outer edge.</param>
    /// <param name="inner">Inner lines; <c>null</c> draws none (existing inner edges stay).</param>
    public void Frame(int row1, int column1, int row2, int column2, XlsxBorder outer, XlsxBorder? inner = null)
    {
        for (var r = row1; r <= row2; r++)
        {
            for (var c = column1; c <= column2; c++)
            {
                ref var s = ref At(r, c).Style;
                if (c == column1) { s.Left = outer; } else if (inner != null) { s.Left = inner.Value; }
                if (c == column2) { s.Right = outer; } else if (inner != null) { s.Right = inner.Value; }
                if (r == row1) { s.Top = outer; } else if (inner != null) { s.Top = inner.Value; }
                if (r == row2) { s.Bottom = outer; } else if (inner != null) { s.Bottom = inner.Value; }
            }
        }
    }

    /// <summary>
    /// Merges a range into one cell; the value stays in the top-left cell.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <exception cref="ArgumentException">The range is reversed.</exception>
    public void Merge(int row1, int column1, int row2, int column2)
    {
        Check(row1, column1);
        Check(row2, column2);
        if (row2 < row1 || column2 < column1)
        {
            throw new ArgumentException(Messages.Get(Strings.MergeRangeReversed));
        }

        if (row1 != row2 || column1 != column2)
        {
            _merges.Add((row1, column1, row2, column2));
        }
    }

    /// <summary>
    /// Sets a column width in Excel units (the number of "0" characters of the default font).
    /// </summary>
    /// <param name="column">Column (1-based).</param>
    /// <param name="width">Width in Excel units.</param>
    public void ColumnWidth(int column, double width)
    {
        Check(1, column);
        _widths[column] = width;
    }

    /// <summary>
    /// Hides a column; its values stay and charts still draw them (for example a helper column).
    /// </summary>
    /// <param name="column">Column (1-based).</param>
    public void HideColumn(int column)
    {
        Check(1, column);
        _hidden.Add(column);
    }

    /// <summary>
    /// Sets a row height in points.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="height">Height in points.</param>
    public void RowHeight(int row, double height)
    {
        Check(row, 1);
        _heights[row] = height;
    }

    /// <summary>
    /// Links a cell to another sheet of the workbook; clicking it goes to A1 of that sheet.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="sheet">Name of the target sheet.</param>
    public void Link(int row, int column, string sheet)
    {
        Check(row, column);
        _links.Add((row, column, sheet));
    }

    /// <summary>
    /// Links a cell to a sheet in another Excel file (a HYPERLINK formula): clicking it opens that file at A1 of the
    /// sheet. The path is relative to this file's folder; the files must move together.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="text">Text shown in the cell.</param>
    /// <param name="file">Path of the target file, relative to this file.</param>
    /// <param name="sheet">Name of the target sheet.</param>
    public void LinkFile(int row, int column, string text, string file, string sheet)
    {
        Check(row, column);
        var cell = At(row, column);
        cell.Value = text;
        var target = file + "#'" + sheet.Replace("'", "''") + "'!A1";
        if (target.Length > 250)
        {
            // The first argument of HYPERLINK holds at most 255 characters; beyond that the cell stays plain text.
            cell.Formula = null;
            return;
        }

        cell.Formula = "HYPERLINK(\"" + target.Replace("\"", "\"\"") + "\",\"" + text.Replace("\"", "\"\"") + "\")";
    }

    /// <summary>
    /// Puts an auto filter on a header row and the range below it (drop-down arrows on the column headers).
    /// </summary>
    /// <param name="row1">Header row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    public void AutoFilter(int row1, int column1, int row2, int column2)
    {
        Check(row1, column1);
        Check(row2, column2);
        AutoFilterRange = (row1, column1, row2, column2);
    }

    /// <summary>
    /// The number in a cell; <c>null</c> when the cell is empty or holds text. Fills the chart caches.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <returns>The number or <c>null</c>.</returns>
    public double? NumberAt(int row, int column) =>
        _cells.TryGetValue(Key(row, column), out var cell) && cell.Value is double v ? v : (double?)null;

    #endregion

    #region Charts

    /// <summary>
    /// Adds a chart to the sheet; add its series afterwards with <see cref="XlsxChart.AddSeries"/>.
    /// </summary>
    /// <param name="row">Row of the top-left corner.</param>
    /// <param name="column">Column of the top-left corner.</param>
    /// <param name="width">Width in pixels.</param>
    /// <param name="height">Height in pixels.</param>
    /// <returns>The new chart.</returns>
    public XlsxChart AddChart(int row, int column, int width, int height)
    {
        Check(row, column);
        var chart = new XlsxChart(this, row, column, width, height);
        _charts.Add(chart);
        return chart;
    }

    #endregion

    #region Internals

    /// <summary>
    /// Cell key: row and column in one number, ordered by row first.
    /// </summary>
    internal static long Key(int row, int column) => (long)row * (MaxColumn + 1) + column;

    internal static int RowOf(long key) => (int)(key / (MaxColumn + 1));

    internal static int ColumnOf(long key) => (int)(key % (MaxColumn + 1));

    private Cell At(int row, int column)
    {
        Check(row, column);
        var key = Key(row, column);
        if (!_cells.TryGetValue(key, out var cell))
        {
            cell = new Cell();
            _cells[key] = cell;
        }

        return cell;
    }

    private static void Check(int row, int column)
    {
        if (row < 1 || row > MaxRow || column < 1 || column > MaxColumn)
        {
            throw new ArgumentOutOfRangeException(nameof(row), Messages.Format(Strings.CellOutOfRange, row, column));
        }
    }

    /// <summary>
    /// An image on the sheet.
    /// </summary>
    internal sealed class XlsxImage
    {
        public XlsxImage(int row, int column, byte[] png, int width, int height)
        {
            Row = row;
            Column = column;
            Png = png;
            Width = width;
            Height = height;
        }

        public int Row { get; }

        public int Column { get; }

        public byte[] Png { get; }

        public int Width { get; }

        public int Height { get; }
    }

    /// <summary>
    /// A colour scale rule.
    /// </summary>
    internal sealed class XlsxColorScale
    {
        public XlsxColorScale(int row1, int column1, int row2, int column2, string low, string? mid, string high)
        {
            Row1 = row1;
            Column1 = column1;
            Row2 = row2;
            Column2 = column2;
            Low = low;
            Mid = mid;
            High = high;
        }

        public int Row1 { get; }

        public int Column1 { get; }

        public int Row2 { get; }

        public int Column2 { get; }

        public string Low { get; }

        public string? Mid { get; }

        public string High { get; }
    }

    /// <summary>
    /// A cell: value (text, number or empty) and style.
    /// </summary>
    internal sealed class Cell
    {
        public object? Value;

        /// <summary>
        /// When set, the cell is a formula (written without the leading <c>=</c>); <see cref="Value"/> is the cached
        /// result text.
        /// </summary>
        public string? Formula;

        public XlsxStyle Style;
    }

    #endregion
}
