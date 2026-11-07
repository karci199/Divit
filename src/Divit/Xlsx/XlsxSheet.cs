#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;

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
    private readonly Dictionary<int, Line> _rows = new Dictionary<int, Line>();
    private readonly Dictionary<int, Line> _columns = new Dictionary<int, Line>();
    private readonly List<(int Row1, int Col1, int Row2, int Col2)> _merges = new List<(int, int, int, int)>();
    private readonly List<(int Row, int Col, string Sheet)> _links = new List<(int, int, string)>();
    private readonly List<(int Row, int Col, string Url)> _urlLinks = new List<(int, int, string)>();
    private readonly SortedDictionary<long, XlsxComment> _comments = new SortedDictionary<long, XlsxComment>();
    private readonly List<XlsxChart> _charts = new List<XlsxChart>();
    private readonly List<XlsxImage> _images = new List<XlsxImage>();
    private readonly List<XlsxConditional> _conditionals = new List<XlsxConditional>();
    private readonly List<XlsxTable> _tables = new List<XlsxTable>();
    private readonly List<(int Row1, int Col1, int Row2, int Col2, XlsxValidation Validation)> _validations =
        new List<(int, int, int, int, XlsxValidation)>();

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

    /// <summary>
    /// Conditional formats (colour scales, highlights, data bars, icon sets) in the order they were added, which is
    /// also their priority.
    /// </summary>
    internal IReadOnlyList<XlsxConditional> Conditionals => _conditionals;

    internal IReadOnlyList<(int Row1, int Col1, int Row2, int Col2, XlsxValidation Validation)> Validations => _validations;

    /// <summary>
    /// Whether the tab is shown; at least one sheet of a workbook must stay visible.
    /// </summary>
    public XlsxSheetVisibility Visibility { get; set; }

    /// <summary>
    /// Paper the sheet is printed on; A4 by default.
    /// </summary>
    public XlsxPaperSize PaperSize { get; set; } = XlsxPaperSize.A4;

    /// <summary>
    /// Page margins in inches; <see cref="XlsxMargins.Default"/> unless set.
    /// </summary>
    public XlsxMargins Margins { get; set; } = XlsxMargins.Default;

    /// <summary>
    /// Centres the printed content between the left and right margins.
    /// </summary>
    public bool PrintCenteredHorizontally { get; set; }

    /// <summary>
    /// Centres the printed content between the top and bottom margins.
    /// </summary>
    public bool PrintCenteredVertically { get; set; }

    /// <summary>
    /// Prints the gridlines.
    /// </summary>
    public bool PrintGridLines { get; set; }

    internal (int Row1, int Col1, int Row2, int Col2)? AutoFilterRange { get; private set; }

    internal IReadOnlyDictionary<long, Cell> Cells => _cells;

    /// <summary>
    /// Rows with a height, hidden or grouped; the key is the row number.
    /// </summary>
    internal IReadOnlyDictionary<int, Line> Rows => _rows;

    /// <summary>
    /// Columns with a width, hidden or grouped; the key is the column number.
    /// </summary>
    internal IReadOnlyDictionary<int, Line> Columns => _columns;

    internal IReadOnlyList<(int Row1, int Col1, int Row2, int Col2)> Merges => _merges;

    internal IReadOnlyList<(int Row, int Col, string Sheet)> Links => _links;

    internal IReadOnlyList<(int Row, int Col, string Url)> UrlLinks => _urlLinks;

    /// <summary>
    /// Comments in row order, then column order.
    /// </summary>
    internal IReadOnlyCollection<XlsxComment> Comments => _comments.Values;

    #endregion

    #region Protection

    /// <summary>
    /// Locks the sheet: locked cells cannot be edited by hand; selecting, copying and clicking links still work. Cells
    /// whose style has <see cref="XlsxStyle.Unlocked"/> stay editable, which turns the sheet into a form.
    /// </summary>
    /// <remarks>
    /// The password is stored as a salted SHA-512 hash with 100 000 rounds, as Excel 2013 and later write it. A lock is
    /// not encryption: the values stay readable by any program; it guards against changes by hand. To prove that a file
    /// was not modified, keep a hash of the whole file elsewhere.
    /// </remarks>
    /// <param name="password">Password; none when empty (Excel then lifts the lock with one click).</param>
    /// <param name="allow">What users may still do on the locked sheet.</param>
    public void Protect(string? password = null, XlsxAllow allow = XlsxAllow.None)
    {
        Protected = true;
        Password = password;
        Allowed = allow;
    }

    /// <summary>
    /// What users may still do on the sheet while it is locked; see <see cref="Protect"/>.
    /// </summary>
    public XlsxAllow Allowed { get; private set; }

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
        if (!OfficeText.IsPng(png))
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
        CheckRange(row1, column1, row2, column2);
        _conditionals.Add(new XlsxColorScale(row1, column1, row2, column2, lowColor, midColor, highColor));
    }

    /// <summary>
    /// Gives the cells of a range for which a rule holds a different look (Excel conditional formatting), for example
    /// a red fill for values below a limit. Rules apply in the order they were added; a later rule on the same cell
    /// adds to the look of an earlier one.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="rule">The condition, from the factories of <see cref="XlsxRule"/>.</param>
    /// <param name="look">The look of matching cells. Font colour, bold, italic, underline, strikethrough, fill,
    /// borders and number format apply; font name, size and alignment do not.</param>
    public void Highlight(int row1, int column1, int row2, int column2, XlsxRule rule, XlsxStyle look)
    {
        CheckRange(row1, column1, row2, column2);
        _conditionals.Add(new XlsxHighlight(row1, column1, row2, column2, rule, look));
    }

    /// <summary>
    /// Draws a bar in every cell of a range, as long as the value is large compared with the others (Excel conditional
    /// formatting). The values stay readable on top of the bars.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="color">Bar colour as <c>#RRGGBB</c>.</param>
    public void DataBar(int row1, int column1, int row2, int column2, string color)
    {
        CheckRange(row1, column1, row2, column2);
        XlsxWorkbook.Argb(color);
        _conditionals.Add(new XlsxDataBar(row1, column1, row2, column2, color));
    }

    /// <summary>
    /// Puts an icon in every cell of a range by where its value falls between the lowest and highest value (Excel
    /// conditional formatting): with three icons, the lowest third, the middle third and the highest third.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="icons">The icon set.</param>
    /// <param name="reverse">Gives the low values the icons of the high ones (for example red for a high error).</param>
    /// <param name="showValue">Whether the value is shown next to the icon.</param>
    public void IconSet(int row1, int column1, int row2, int column2, XlsxIconSet icons, bool reverse = false, bool showValue = true)
    {
        CheckRange(row1, column1, row2, column2);
        _conditionals.Add(new XlsxIcons(row1, column1, row2, column2, icons, reverse, showValue));
    }

    /// <summary>
    /// Limits what may be typed into a range: a drop-down list, numbers or dates within limits, a text length or a
    /// formula. A later validation on the same cells wins in Excel; do not overlap them.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="validation">What is allowed, from the factories of <see cref="XlsxValidation"/>.</param>
    public void Validate(int row1, int column1, int row2, int column2, XlsxValidation validation)
    {
        CheckRange(row1, column1, row2, column2);
        _validations.Add((row1, column1, row2, column2, validation));
    }

    /// <summary>
    /// Writes a formula; Excel computes it when the file opens (for example <c>SUM(B2:B10)</c> or <c>Sales*Rate</c>).
    /// Formulas use English function names and commas, with or without a leading <c>=</c>.
    /// </summary>
    /// <remarks>
    /// Without a cached value, programs that do not compute on opening (file previews, some viewers) show an empty
    /// cell; give the expected result with the other overloads when it is known.
    /// </remarks>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="formula">The formula, at most 8192 characters.</param>
    /// <exception cref="ArgumentException">The formula is empty or too long.</exception>
    public void SetFormula(int row, int column, string formula) => SetFormula(row, column, formula, (object?)null);

    /// <summary>
    /// Writes a formula with the number it results in, shown until the formula is computed again.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="formula">The formula.</param>
    /// <param name="result">The result.</param>
    public void SetFormula(int row, int column, string formula, double result) =>
        SetFormula(row, column, formula, double.IsNaN(result) || double.IsInfinity(result) ? null : (object)result);

    /// <summary>
    /// Writes a formula with the text it results in, shown until the formula is computed again.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="formula">The formula.</param>
    /// <param name="result">The result.</param>
    public void SetFormula(int row, int column, string formula, string result) => SetFormula(row, column, formula, (object)result);

    /// <summary>
    /// Writes text whose pieces have their own fonts: a bold word, a red warning, a unit in superscript (m²).
    /// </summary>
    /// <example>
    /// <code>
    /// sheet.SetRichText(1, 1, "Status: ", new XlsxRun("late", new XlsxStyle { Bold = true, FontColor = "#C00000" }));
    /// </code>
    /// </example>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="runs">The pieces in order; plain strings take the cell's font. Empty pieces are left out.</param>
    public void SetRichText(int row, int column, params XlsxRun[] runs)
    {
        var cell = At(row, column);
        var pieces = new List<XlsxRun>(runs.Length);
        foreach (var run in runs)
        {
            if (!string.IsNullOrEmpty(run.Text))
            {
                pieces.Add(run);
            }
        }

        cell.Formula = null;
        if (pieces.Count == 0)
        {
            cell.Value = null;
        }
        else if (pieces.Count == 1 && pieces[0].Look == null)
        {
            // One piece in the cell's font is plain text.
            cell.Value = pieces[0].Text;
        }
        else
        {
            cell.Value = new XlsxRichText(pieces);
        }
    }

    /// <summary>
    /// A formula without its leading <c>=</c> and surrounding spaces.
    /// </summary>
    /// <exception cref="ArgumentException">The formula is empty or longer than Excel's 8192 characters.</exception>
    internal static string TrimFormula(string formula)
    {
        var text = formula.Trim();
        if (text.StartsWith("=", StringComparison.Ordinal))
        {
            text = text.Substring(1).TrimStart();
        }

        if (text.Length == 0 || text.Length > 8192)
        {
            throw new ArgumentException(Messages.Get(Strings.FormulaLength), nameof(formula));
        }

        return text;
    }

    private void SetFormula(int row, int column, string formula, object? result)
    {
        var text = TrimFormula(formula);
        var cell = At(row, column);
        cell.Formula = text;
        cell.Value = result;
    }

    #endregion

    #region Cells

    /// <summary>
    /// Writes text into a cell.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="text">The text.</param>
    public void Set(int row, int column, string text) => Value(row, column, text);

    /// <summary>
    /// Writes a number into a cell. NaN and infinity are not numbers in Excel; the cell stays empty.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="number">The number.</param>
    public void Set(int row, int column, double number) =>
        Value(row, column, double.IsNaN(number) || double.IsInfinity(number) ? null : (object)number);

    /// <summary>
    /// Writes a date (and time) into a cell. Excel stores it as a day number; unless the cell already has a number
    /// format, <c>yyyy-mm-dd</c> is used, or <c>yyyy-mm-dd hh:mm:ss</c> when the value has a time of day.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="date">The date; its <see cref="DateTime.Kind"/> is ignored, the clock time is written as it is.</param>
    /// <exception cref="ArgumentOutOfRangeException">The date is before 1900-01-01, Excel's first day.</exception>
    public void Set(int row, int column, DateTime date)
    {
        var cell = At(row, column);
        cell.Value = DateNumber(date);
        cell.Formula = null;
        if (cell.Style.NumberFormat == null)
        {
            cell.Style.NumberFormat = date.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss";
        }
    }

    /// <summary>
    /// Writes TRUE or FALSE into a cell.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="value">The value.</param>
    public void Set(int row, int column, bool value) => Value(row, column, value);

    /// <summary>
    /// Excel's day number of a date: 1 is 1900-01-01. Excel counts a 29 February 1900 that never was, so from March 1900
    /// on the number is one higher than the real day count.
    /// </summary>
    /// <param name="date">The date.</param>
    /// <returns>The day number; the time of day is the fraction.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The date is before 1900-01-01.</exception>
    public static double DateNumber(DateTime date)
    {
        if (date < new DateTime(1900, 1, 1))
        {
            throw new ArgumentOutOfRangeException(nameof(date), Messages.Format(Strings.DateOutOfRange, date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)));
        }

        var days = (date - new DateTime(1899, 12, 30)).Ticks / (double)TimeSpan.TicksPerDay;
        return days < 61 ? days - 1 : days;
    }

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
    /// Gives a range one of Excel's built-in cell styles (Heading 1, Good, Input, Total …): its font, fill and borders,
    /// and the link to the style in Excel's gallery. Number format, alignment and protection of the cells stay.
    /// </summary>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="cellStyle">The built-in style.</param>
    public void ApplyCellStyle(int row1, int column1, int row2, int column2, XlsxCellStyle cellStyle)
    {
        var look = XlsxStyle.Preset(cellStyle);
        Format(row1, column1, row2, column2, (ref XlsxStyle s) =>
        {
            s.FontName = look.FontName;
            s.FontSize = look.FontSize;
            s.Bold = look.Bold;
            s.Italic = look.Italic;
            s.Underline = look.Underline;
            s.DoubleUnderline = look.DoubleUnderline;
            s.Strikethrough = look.Strikethrough;
            s.Script = look.Script;
            s.FontColor = look.FontColor;
            s.Fill = look.Fill;
            s.Left = look.Left;
            s.Right = look.Right;
            s.Top = look.Top;
            s.Bottom = look.Bottom;
            s.CellStyle = cellStyle;
        });
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
        ColumnLine(column).Size = width;
    }

    /// <summary>
    /// Hides a column; its values stay and charts still draw them (for example a helper column).
    /// </summary>
    /// <param name="column">Column (1-based).</param>
    public void HideColumn(int column)
    {
        Check(1, column);
        ColumnLine(column).Hidden = true;
    }

    /// <summary>
    /// Sets a row height in points.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="height">Height in points.</param>
    public void RowHeight(int row, double height)
    {
        Check(row, 1);
        RowLine(row).Size = height;
    }

    /// <summary>
    /// Hides a row; its values stay and formulas and charts still use them.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    public void HideRow(int row)
    {
        Check(row, 1);
        RowLine(row).Hidden = true;
    }

    /// <summary>
    /// Groups rows under an outline button, so they can be folded away. Groups nest: grouping rows that are already
    /// grouped adds a level, up to seven. The button sits on the row below the group (the summary row).
    /// </summary>
    /// <param name="first">First row of the group (1-based).</param>
    /// <param name="last">Last row of the group.</param>
    /// <param name="collapsed">Whether the group opens folded: its rows hidden, the button showing "+".</param>
    public void GroupRows(int first, int last, bool collapsed = false)
    {
        Check(first, 1);
        Check(last, 1);
        Group(_rows, first, last, collapsed, last < MaxRow);
    }

    /// <summary>
    /// Groups columns under an outline button; see <see cref="GroupRows"/>. The button sits on the column to the right
    /// of the group.
    /// </summary>
    /// <param name="first">First column of the group (1-based).</param>
    /// <param name="last">Last column of the group.</param>
    /// <param name="collapsed">Whether the group opens folded.</param>
    public void GroupColumns(int first, int last, bool collapsed = false)
    {
        Check(1, first);
        Check(1, last);
        Group(_columns, first, last, collapsed, last < MaxColumn);
    }

    /// <summary>
    /// Writes a note on a cell: a red corner on the cell, the text in a yellow box when the pointer rests on it.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="text">Text of the note; line breaks are kept.</param>
    /// <param name="author">Author shown in Excel; the workbook author when empty.</param>
    public void Comment(int row, int column, string text, string? author = null)
    {
        Check(row, column);
        _comments[Key(row, column)] = new XlsxComment(row, column, text, author);
    }

    /// <summary>
    /// Links a cell to a web page, an e-mail address (<c>mailto:</c>) or a file; clicking the cell opens it. The cell
    /// shows <paramref name="text"/> and gets the usual link look (blue, underlined) unless it already has a font colour.
    /// </summary>
    /// <param name="row">Row (1-based).</param>
    /// <param name="column">Column (1-based).</param>
    /// <param name="url">Target, for example <c>https://example.com</c>, <c>mailto:someone@example.com</c> or a
    /// relative file path such as <c>report.pdf</c>.</param>
    /// <param name="text">Text shown in the cell; the address itself when empty.</param>
    /// <exception cref="ArgumentException">The address is empty.</exception>
    public void LinkUrl(int row, int column, string url, string? text = null)
    {
        Check(row, column);
        if (url.Length == 0)
        {
            throw new ArgumentException(Messages.Get(Strings.LinkEmpty), nameof(url));
        }

        var cell = At(row, column);
        cell.Value = text ?? url;
        cell.Formula = null;
        if (cell.Style.FontColor == null)
        {
            cell.Style.FontColor = "#0563C1";
            cell.Style.Underline = true;
        }

        _urlLinks.Add((row, column, url));
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

    /// <summary>
    /// Makes a range an Excel table: the first row holds the column names, the rows below the data. Tables get
    /// filter buttons, banded rows and a name formulas can use (<c>SUM(Sales[Amount])</c>).
    /// </summary>
    /// <param name="row1">Header row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last data row; at least one row below the header.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="name">Name of the table, by the rules of defined names (<see cref="XlsxWorkbook.DefineName"/>),
    /// unique in the workbook among tables and names.</param>
    /// <returns>The table, to set its style and totals row.</returns>
    /// <exception cref="ArgumentException">The name is not usable, the range has no data row, or it overlaps another
    /// table on this sheet.</exception>
    public XlsxTable AddTable(int row1, int column1, int row2, int column2, string name)
    {
        CheckRange(row1, column1, row2, column2);
        if (row2 <= row1)
        {
            throw new ArgumentException(Messages.Get(Strings.TableNeedsData));
        }

        if (!XlsxWorkbook.IsDefinedName(name))
        {
            throw new ArgumentException(Messages.Format(Strings.DefinedNameInvalid, name), nameof(name));
        }

        foreach (var other in _tables)
        {
            if (row1 <= other.LastRow && other.Row1 <= row2 && column1 <= other.Column2 && other.Column1 <= column2)
            {
                throw new ArgumentException(Messages.Format(Strings.TablesOverlap, name, other.Name));
            }
        }

        var table = new XlsxTable(this, row1, column1, row2, column2, name);
        _tables.Add(table);
        return table;
    }

    /// <summary>
    /// The tables of the sheet in the order they were added.
    /// </summary>
    public IReadOnlyList<XlsxTable> Tables => _tables;

    /// <summary>
    /// The text of a cell as Excel shows a header: text as it is, numbers in invariant form; <c>null</c> when empty.
    /// </summary>
    internal string? TextAt(int row, int column)
    {
        if (!_cells.TryGetValue(Key(row, column), out var cell))
        {
            return null;
        }

        switch (cell.Value)
        {
            case string text: return text;
            case XlsxRichText rich: return rich.Text;
            case double number: return number.ToString("R", CultureInfo.InvariantCulture);
            case bool flag: return flag ? "TRUE" : "FALSE";
            default: return null;
        }
    }

    /// <summary>
    /// The number format of a cell; <c>null</c> for General.
    /// </summary>
    internal string? FormatAt(int row, int column) =>
        _cells.TryGetValue(Key(row, column), out var cell) ? cell.Style.NumberFormat : null;

    /// <summary>
    /// Whether a cell holds plain text without a formula, as a table header must.
    /// </summary>
    internal bool IsPlainText(int row, int column) =>
        _cells.TryGetValue(Key(row, column), out var cell) && cell.Value is string && cell.Formula == null;

    /// <summary>
    /// Whether a cell holds a value or a formula.
    /// </summary>
    internal bool HasValue(int row, int column) =>
        _cells.TryGetValue(Key(row, column), out var cell) && (cell.Value != null || cell.Formula != null);

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

    /// <summary>
    /// Writes a value; a formula in the cell goes, the style stays.
    /// </summary>
    private void Value(int row, int column, object? value)
    {
        var cell = At(row, column);
        cell.Value = value;
        cell.Formula = null;
    }

    private Line RowLine(int row) => LineOf(_rows, row);

    private Line ColumnLine(int column) => LineOf(_columns, column);

    private static Line LineOf(Dictionary<int, Line> lines, int number)
    {
        if (!lines.TryGetValue(number, out var line))
        {
            line = new Line();
            lines[number] = line;
        }

        return line;
    }

    private static void Group(Dictionary<int, Line> lines, int first, int last, bool collapsed, bool hasSummary)
    {
        if (last < first)
        {
            throw new ArgumentException(Messages.Get(Strings.GroupRangeReversed));
        }

        for (var i = first; i <= last; i++)
        {
            var line = LineOf(lines, i);
            line.Level = Math.Min(7, line.Level + 1);
            line.Hidden |= collapsed;
        }

        // Excel keeps the folded state on the row or column after the group.
        if (collapsed && hasSummary)
        {
            LineOf(lines, last + 1).Collapsed = true;
        }
    }

    private static void CheckRange(int row1, int column1, int row2, int column2)
    {
        Check(row1, column1);
        Check(row2, column2);
        if (row2 < row1 || column2 < column1)
        {
            throw new ArgumentException(Messages.Get(Strings.RangeReversed));
        }
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
    /// Properties of a row or a column: size, hidden, outline level.
    /// </summary>
    internal sealed class Line
    {
        /// <summary>
        /// Height in points (row) or width in Excel units (column); <c>null</c> for the default.
        /// </summary>
        public double? Size;

        public bool Hidden;

        /// <summary>
        /// Outline level, 0 when not grouped.
        /// </summary>
        public int Level;

        /// <summary>
        /// Whether the group before this row or column is folded.
        /// </summary>
        public bool Collapsed;
    }

    /// <summary>
    /// A note on a cell.
    /// </summary>
    internal sealed class XlsxComment
    {
        public XlsxComment(int row, int column, string text, string? author)
        {
            Row = row;
            Column = column;
            Text = text;
            Author = author;
        }

        public int Row { get; }

        public int Column { get; }

        public string Text { get; }

        public string? Author { get; }
    }

    /// <summary>
    /// A conditional format on a range.
    /// </summary>
    internal abstract class XlsxConditional
    {
        protected XlsxConditional(int row1, int column1, int row2, int column2)
        {
            Row1 = row1;
            Column1 = column1;
            Row2 = row2;
            Column2 = column2;
        }

        public int Row1 { get; }

        public int Column1 { get; }

        public int Row2 { get; }

        public int Column2 { get; }
    }

    /// <summary>
    /// A rule with a look: the cells for which the rule holds take the look.
    /// </summary>
    internal sealed class XlsxHighlight : XlsxConditional
    {
        public XlsxHighlight(int row1, int column1, int row2, int column2, XlsxRule rule, XlsxStyle look)
            : base(row1, column1, row2, column2)
        {
            Rule = rule;
            Look = look;
        }

        public XlsxRule Rule { get; }

        public XlsxStyle Look { get; }
    }

    /// <summary>
    /// Bars in the cells, as long as the value.
    /// </summary>
    internal sealed class XlsxDataBar : XlsxConditional
    {
        public XlsxDataBar(int row1, int column1, int row2, int column2, string color)
            : base(row1, column1, row2, column2)
        {
            Color = color;
        }

        public string Color { get; }
    }

    /// <summary>
    /// Icons in the cells by value.
    /// </summary>
    internal sealed class XlsxIcons : XlsxConditional
    {
        public XlsxIcons(int row1, int column1, int row2, int column2, XlsxIconSet set, bool reverse, bool showValue)
            : base(row1, column1, row2, column2)
        {
            Set = set;
            Reverse = reverse;
            ShowValue = showValue;
        }

        public XlsxIconSet Set { get; }

        public bool Reverse { get; }

        public bool ShowValue { get; }
    }

    /// <summary>
    /// A colour scale rule.
    /// </summary>
    internal sealed class XlsxColorScale : XlsxConditional
    {
        public XlsxColorScale(int row1, int column1, int row2, int column2, string low, string? mid, string high)
            : base(row1, column1, row2, column2)
        {
            Low = low;
            Mid = mid;
            High = high;
        }

        public string Low { get; }

        public string? Mid { get; }

        public string High { get; }
    }

    /// <summary>
    /// A cell: value (text, number, boolean or empty), formula and style.
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
