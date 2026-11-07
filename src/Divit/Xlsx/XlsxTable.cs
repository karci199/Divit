#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;

/// <summary>
/// The function of a table column in the totals row.
/// </summary>
public enum XlsxTotal
{
    /// <summary>
    /// No total.
    /// </summary>
    None,

    /// <summary>
    /// The sum.
    /// </summary>
    Sum,

    /// <summary>
    /// The average of the numbers.
    /// </summary>
    Average,

    /// <summary>
    /// How many cells are not empty.
    /// </summary>
    Count,

    /// <summary>
    /// How many cells hold numbers.
    /// </summary>
    CountNumbers,

    /// <summary>
    /// The smallest number.
    /// </summary>
    Min,

    /// <summary>
    /// The largest number.
    /// </summary>
    Max,

    /// <summary>
    /// The standard deviation of a sample.
    /// </summary>
    StdDev,

    /// <summary>
    /// The variance of a sample.
    /// </summary>
    Var,
}

/// <summary>
/// An Excel table (ListObject) on a range: a header row with filter buttons, banded rows in a table style, an
/// optional totals row, and a name that formulas use (<c>SUM(Sales[Amount])</c>). Add it with
/// <see cref="XlsxSheet.AddTable"/>.
/// </summary>
/// <remarks>
/// The first row of the range holds the column names. When the workbook is saved, an empty header cell gets
/// <c>Column1</c>, <c>Column2</c> …, a repeated name gets a number (Excel needs unique names), and the totals row is
/// written below the last row of the range.
/// </remarks>
public sealed class XlsxTable
{
    private readonly XlsxSheet _sheet;
    private readonly XlsxTotal[] _totals;
    private readonly string?[] _labels;
    private string? _style = "TableStyleMedium2";

    internal XlsxTable(XlsxSheet sheet, int row1, int column1, int row2, int column2, string name)
    {
        _sheet = sheet;
        Row1 = row1;
        Column1 = column1;
        Row2 = row2;
        Column2 = column2;
        Name = name;
        _totals = new XlsxTotal[column2 - column1 + 1];
        _labels = new string?[column2 - column1 + 1];
        ColumnNames = Array.Empty<string>();
    }

    /// <summary>
    /// The table's name, used in formulas and in Excel's Name Box.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// The table style, one of Excel's: <c>TableStyleLight1</c>–<c>21</c>, <c>TableStyleMedium1</c>–<c>28</c>,
    /// <c>TableStyleDark1</c>–<c>11</c>; <c>null</c> for none.
    /// </summary>
    /// <exception cref="ArgumentException">Not one of Excel's table styles.</exception>
    public string? Style
    {
        get => _style;
        set
        {
            if (value != null && !IsTableStyle(value))
            {
                throw new ArgumentException(Messages.Format(Strings.TableStyleUnknown, value), nameof(value));
            }

            _style = value;
        }
    }

    /// <summary>
    /// Shades every other row.
    /// </summary>
    public bool RowStripes { get; set; } = true;

    /// <summary>
    /// Shades every other column.
    /// </summary>
    public bool ColumnStripes { get; set; }

    /// <summary>
    /// Emphasises the first column.
    /// </summary>
    public bool FirstColumn { get; set; }

    /// <summary>
    /// Emphasises the last column.
    /// </summary>
    public bool LastColumn { get; set; }

    /// <summary>
    /// Shows the filter buttons on the header row.
    /// </summary>
    public bool FilterButtons { get; set; } = true;

    /// <summary>
    /// Adds a totals row below the last row of the range; give its functions with <see cref="Total"/>.
    /// </summary>
    public bool TotalsRow { get; set; }

    /// <summary>
    /// The column names as written by the last save.
    /// </summary>
    public IReadOnlyList<string> ColumnNames { get; private set; }

    internal int Row1 { get; }

    internal int Column1 { get; }

    internal int Row2 { get; }

    internal int Column2 { get; }

    /// <summary>
    /// Last row of the table, the totals row included.
    /// </summary>
    internal int LastRow => TotalsRow ? Row2 + 1 : Row2;

    /// <summary>
    /// Number of the table in the workbook, from 1; set when saving.
    /// </summary>
    internal int Id { get; set; }

    /// <summary>
    /// Gives a column of the totals row a function (<c>SUBTOTAL</c>, which skips rows hidden by the filter) and
    /// turns the totals row on.
    /// </summary>
    /// <param name="column">Column within the table, 1 for its first column.</param>
    /// <param name="function">The function.</param>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the table.</exception>
    public void Total(int column, XlsxTotal function)
    {
        CheckColumn(column);
        _totals[column - 1] = function;
        _labels[column - 1] = null;
        TotalsRow = true;
    }

    /// <summary>
    /// Writes a text into a column of the totals row (for example "Total") and turns the totals row on.
    /// </summary>
    /// <param name="column">Column within the table, 1 for its first column.</param>
    /// <param name="text">The text.</param>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the table.</exception>
    public void TotalLabel(int column, string text)
    {
        CheckColumn(column);
        _labels[column - 1] = text;
        _totals[column - 1] = XlsxTotal.None;
        TotalsRow = true;
    }

    internal XlsxTotal TotalOf(int index) => _totals[index];

    internal string? LabelOf(int index) => _labels[index];

    /// <summary>
    /// Fixes the column names from the header cells, writing names into empty or repeated ones, and fills the
    /// totals row: labels as text, functions as <c>SUBTOTAL</c> formulas with their current result.
    /// </summary>
    internal void Prepare()
    {
        var names = new List<string>();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var c = Column1; c <= Column2; c++)
        {
            var text = _sheet.TextAt(Row1, c);
            var name = string.IsNullOrEmpty(text) ? "Column" + (c - Column1 + 1).ToString(CultureInfo.InvariantCulture) : text!;
            var unique = name;
            for (var n = 2; !used.Add(unique); n++)
            {
                unique = name + n.ToString(CultureInfo.InvariantCulture);
            }

            if (unique != text || !_sheet.IsPlainText(Row1, c))
            {
                // A header must hold its name as plain text: numbers, rich text and formulas become the name.
                _sheet.Set(Row1, c, unique);
            }

            names.Add(unique);
        }

        ColumnNames = names;
        if (!TotalsRow)
        {
            return;
        }

        for (var i = 0; i < names.Count; i++)
        {
            var column = Column1 + i;
            if (_labels[i] != null)
            {
                _sheet.Set(Row2 + 1, column, _labels[i]!);
            }
            else if (_totals[i] != XlsxTotal.None)
            {
                var formula = "SUBTOTAL(" + Code(_totals[i]).ToString(CultureInfo.InvariantCulture) + "," + Name + "[" + Escape(names[i]) + "])";
                var result = Result(_totals[i], column);
                if (result is double value)
                {
                    _sheet.SetFormula(Row2 + 1, column, formula, value);
                }
                else
                {
                    _sheet.SetFormula(Row2 + 1, column, formula);
                }
            }
        }
    }

    /// <summary>
    /// The totalsRowFunction of a total.
    /// </summary>
    internal static string FunctionName(XlsxTotal total)
    {
        switch (total)
        {
            case XlsxTotal.Sum: return "sum";
            case XlsxTotal.Average: return "average";
            case XlsxTotal.Count: return "count";
            case XlsxTotal.CountNumbers: return "countNums";
            case XlsxTotal.Min: return "min";
            case XlsxTotal.Max: return "max";
            case XlsxTotal.StdDev: return "stdDev";
            default: return "var";
        }
    }

    /// <summary>
    /// A column name in a structured reference: <c>[</c>, <c>]</c>, <c>#</c> and <c>'</c> take a <c>'</c> before them.
    /// </summary>
    internal static string Escape(string name)
    {
        var escaped = new System.Text.StringBuilder(name.Length + 4);
        foreach (var ch in name)
        {
            if (ch == '[' || ch == ']' || ch == '#' || ch == '\'')
            {
                escaped.Append('\'');
            }

            escaped.Append(ch);
        }

        return escaped.ToString();
    }

    /// <summary>
    /// The SUBTOTAL function number; 100 and above skip rows hidden by hand as well as by the filter.
    /// </summary>
    private static int Code(XlsxTotal total)
    {
        switch (total)
        {
            case XlsxTotal.Average: return 101;
            case XlsxTotal.CountNumbers: return 102;
            case XlsxTotal.Count: return 103;
            case XlsxTotal.Max: return 104;
            case XlsxTotal.Min: return 105;
            case XlsxTotal.StdDev: return 107;
            case XlsxTotal.Sum: return 109;
            default: return 110;
        }
    }

    private static bool IsTableStyle(string style)
    {
        foreach (var (prefix, last) in new[] { ("TableStyleLight", 21), ("TableStyleMedium", 28), ("TableStyleDark", 11) })
        {
            if (style.StartsWith(prefix, StringComparison.Ordinal)
                && int.TryParse(style.Substring(prefix.Length), NumberStyles.None, CultureInfo.InvariantCulture, out var number)
                && number >= 1 && number <= last && style == prefix + number.ToString(CultureInfo.InvariantCulture))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The result Excel will compute, so viewers that do not compute show it too; <c>null</c> when Excel shows an
    /// error (average or deviation of too few numbers).
    /// </summary>
    private double? Result(XlsxTotal total, int column)
    {
        var numbers = new List<double>();
        var filled = 0;
        for (var r = Row1 + 1; r <= Row2; r++)
        {
            if (_sheet.HasValue(r, column))
            {
                filled++;
            }

            if (_sheet.NumberAt(r, column) is double value)
            {
                numbers.Add(value);
            }
        }

        double sum = 0;
        foreach (var value in numbers)
        {
            sum += value;
        }

        switch (total)
        {
            case XlsxTotal.Sum: return sum;
            case XlsxTotal.Count: return filled;
            case XlsxTotal.CountNumbers: return numbers.Count;
            case XlsxTotal.Average: return numbers.Count == 0 ? (double?)null : sum / numbers.Count;
            case XlsxTotal.Min: return numbers.Count == 0 ? 0 : Min(numbers);
            case XlsxTotal.Max: return numbers.Count == 0 ? 0 : Max(numbers);
            default:
                if (numbers.Count < 2)
                {
                    return null;
                }

                var mean = sum / numbers.Count;
                double squares = 0;
                foreach (var value in numbers)
                {
                    squares += (value - mean) * (value - mean);
                }

                var variance = squares / (numbers.Count - 1);
                return total == XlsxTotal.StdDev ? Math.Sqrt(variance) : variance;
        }
    }

    private static double Min(List<double> values)
    {
        var min = values[0];
        foreach (var value in values)
        {
            min = Math.Min(min, value);
        }

        return min;
    }

    private static double Max(List<double> values)
    {
        var max = values[0];
        foreach (var value in values)
        {
            max = Math.Max(max, value);
        }

        return max;
    }

    private void CheckColumn(int column)
    {
        if (column < 1 || column > Column2 - Column1 + 1)
        {
            throw new ArgumentOutOfRangeException(nameof(column), Messages.Format(Strings.TableColumnOutOfRange, column, Name));
        }
    }
}
