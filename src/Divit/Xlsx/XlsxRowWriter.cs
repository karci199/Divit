#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;

/// <summary>
/// Writes the rows of a streamed sheet (<see cref="XlsxWorkbook.AddStreamedSheet"/>) straight into the file while it
/// is saved: rows go out one after another and are not kept, so memory stays the same for a thousand rows or a
/// million.
/// </summary>
/// <remarks>
/// <para>Rows run top to bottom and cells left to right; what is written cannot be changed. A row starts with the
/// first cell (or <see cref="StartRow"/>, to give it a height) and ends with <see cref="EndRow"/>.</para>
/// <example>
/// <code>
/// book.AddStreamedSheet("Log", rows =>
/// {
///     rows.Cell("Time", bold).Cell("Value", bold).EndRow();
///     foreach (var m in measurements)
///     {
///         rows.Cell(m.Time).Cell(m.Value).EndRow();
///     }
/// });
/// </code>
/// </example>
/// </remarks>
public sealed class XlsxRowWriter
{
    private readonly XlsxRowSink _sink;
    private bool _open;

    internal XlsxRowWriter(XlsxRowSink sink)
    {
        _sink = sink;
        Row = 1;
        Column = 1;
    }

    /// <summary>
    /// Number of the row being written, or of the next one (1-based).
    /// </summary>
    public int Row { get; private set; }

    /// <summary>
    /// Column the next cell goes into (1-based).
    /// </summary>
    public int Column { get; private set; }

    /// <summary>
    /// Starts a row with its own height or outline; without it the first cell starts the row.
    /// </summary>
    /// <param name="height">Height in points; the default when empty.</param>
    /// <param name="hidden">Whether the row is hidden.</param>
    /// <param name="outlineLevel">Outline (group) level, 0–7.</param>
    /// <exception cref="InvalidOperationException">A row is already open.</exception>
    public void StartRow(double? height = null, bool hidden = false, int outlineLevel = 0)
    {
        if (_open)
        {
            throw new InvalidOperationException(Messages.Get(Strings.RowAlreadyStarted));
        }

        CheckRow();
        _sink.StartRow(Row, height, hidden, Math.Max(0, Math.Min(7, outlineLevel)));
        _open = true;
    }

    /// <summary>
    /// Ends the row; the next cell goes into column 1 of the next row.
    /// </summary>
    public void EndRow()
    {
        if (_open)
        {
            _sink.EndRow();
            _open = false;
        }

        Row++;
        Column = 1;
    }

    /// <summary>
    /// Leaves rows empty; ends the open row first.
    /// </summary>
    /// <param name="count">Number of empty rows.</param>
    public void SkipRows(int count)
    {
        if (_open || Column > 1)
        {
            EndRow();
        }

        Row += Math.Max(0, count);
    }

    /// <summary>
    /// Leaves cells empty.
    /// </summary>
    /// <param name="count">Number of empty cells.</param>
    /// <returns>This writer, for the next cell.</returns>
    public XlsxRowWriter Skip(int count = 1)
    {
        Column += Math.Max(0, count);
        return this;
    }

    /// <summary>
    /// Writes a text cell; <c>null</c> leaves the cell empty (with the style, when one is given).
    /// </summary>
    /// <param name="text">The text.</param>
    /// <param name="style">The cell's style.</param>
    /// <returns>This writer, for the next cell.</returns>
    public XlsxRowWriter Cell(string? text, in XlsxStyle style = default)
    {
        Begin();
        if (text == null)
        {
            _sink.Empty(Row, Column, style);
        }
        else
        {
            _sink.Text(Row, Column, text, style);
        }

        Column++;
        return this;
    }

    /// <summary>
    /// Writes a number cell; NaN and infinity leave the cell empty.
    /// </summary>
    /// <param name="number">The number.</param>
    /// <param name="style">The cell's style.</param>
    /// <returns>This writer, for the next cell.</returns>
    public XlsxRowWriter Cell(double number, in XlsxStyle style = default)
    {
        Begin();
        if (double.IsNaN(number) || double.IsInfinity(number))
        {
            _sink.Empty(Row, Column, style);
        }
        else
        {
            _sink.Number(Row, Column, number, style);
        }

        Column++;
        return this;
    }

    /// <summary>
    /// Writes a date cell; without a number format in the style it shows as <c>yyyy-mm-dd</c>, or with the time when
    /// it has one (as <see cref="XlsxSheet.Set(int, int, DateTime)"/>).
    /// </summary>
    /// <param name="date">The date.</param>
    /// <param name="style">The cell's style.</param>
    /// <returns>This writer, for the next cell.</returns>
    /// <exception cref="ArgumentOutOfRangeException">The date is before 1900-01-01.</exception>
    public XlsxRowWriter Cell(DateTime date, in XlsxStyle style = default)
    {
        Begin();
        var number = XlsxSheet.DateNumber(date);
        if (style.NumberFormat != null)
        {
            _sink.Number(Row, Column, number, style);
        }
        else
        {
            var dated = style;
            dated.NumberFormat = date.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss";
            _sink.Number(Row, Column, number, dated);
        }

        Column++;
        return this;
    }

    /// <summary>
    /// Writes TRUE or FALSE.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="style">The cell's style.</param>
    /// <returns>This writer, for the next cell.</returns>
    public XlsxRowWriter Cell(bool value, in XlsxStyle style = default)
    {
        Begin();
        _sink.Bool(Row, Column, value, style);
        Column++;
        return this;
    }

    /// <summary>
    /// Writes a cell of pieces with their own fonts (see <see cref="XlsxSheet.SetRichText"/>).
    /// </summary>
    /// <param name="runs">The pieces.</param>
    /// <returns>This writer, for the next cell.</returns>
    public XlsxRowWriter RichText(params XlsxRun[] runs)
    {
        Begin();
        var pieces = new System.Collections.Generic.List<XlsxRun>(runs.Length);
        foreach (var run in runs)
        {
            if (!string.IsNullOrEmpty(run.Text))
            {
                pieces.Add(run);
            }
        }

        if (pieces.Count > 0)
        {
            _sink.Rich(Row, Column, new XlsxRichText(pieces));
        }

        Column++;
        return this;
    }

    /// <summary>
    /// Writes a formula (see <see cref="XlsxSheet.SetFormula(int, int, string)"/>); Excel computes it on opening.
    /// </summary>
    /// <param name="formula">The formula, with or without the leading <c>=</c>.</param>
    /// <param name="result">The result to show until it is computed; none when empty.</param>
    /// <param name="style">The cell's style.</param>
    /// <returns>This writer, for the next cell.</returns>
    /// <exception cref="ArgumentException">The formula is empty or too long.</exception>
    public XlsxRowWriter Formula(string formula, double? result = null, in XlsxStyle style = default)
    {
        var text = XlsxSheet.TrimFormula(formula);
        Begin();
        _sink.Formula(Row, Column, text, result, style);
        Column++;
        return this;
    }

    /// <summary>
    /// Writes a row of values from column 1 and ends it: text, numbers (any numeric type), dates, booleans; <c>null</c>
    /// leaves a cell empty. Handy, but every value is boxed; the <see cref="Cell(double, in XlsxStyle)"/> calls are
    /// the faster way for millions of cells.
    /// </summary>
    /// <param name="values">The values.</param>
    /// <exception cref="ArgumentException">A value is of a type a cell cannot hold.</exception>
    public void WriteRow(params object?[] values)
    {
        foreach (var value in values)
        {
            switch (value)
            {
                case null: Skip(); break;
                case string text: Cell(text); break;
                case DateTime date: Cell(date); break;
                case bool flag: Cell(flag); break;
                case double d: Cell(d); break;
                case float f: Cell(f); break;
                case int i: Cell(i); break;
                case long l: Cell(l); break;
                case decimal m: Cell((double)m); break;
                case short s: Cell(s); break;
                case byte b: Cell(b); break;
                default: throw new ArgumentException(Messages.Format(Strings.CellValueType, value.GetType().Name), nameof(values));
            }
        }

        EndRow();
    }

    /// <summary>
    /// Closes an open row at the end of the callback.
    /// </summary>
    internal void Finish()
    {
        if (_open)
        {
            _sink.EndRow();
            _open = false;
        }
    }

    private void Begin()
    {
        if (Column > XlsxSheet.MaxColumn)
        {
            throw new ArgumentOutOfRangeException(nameof(Column), Messages.Format(Strings.CellOutOfRange, Row, Column));
        }

        if (!_open)
        {
            StartRow();
        }
    }

    private void CheckRow()
    {
        if (Row > XlsxSheet.MaxRow)
        {
            throw new ArgumentOutOfRangeException(nameof(Row), Messages.Format(Strings.CellOutOfRange, Row, Column));
        }
    }
}

/// <summary>
/// Where a row writer's cells go: the sheet part being written.
/// </summary>
internal abstract class XlsxRowSink
{
    public abstract void StartRow(int row, double? height, bool hidden, int outlineLevel);

    public abstract void EndRow();

    public abstract void Empty(int row, int column, in XlsxStyle style);

    public abstract void Text(int row, int column, string text, in XlsxStyle style);

    public abstract void Number(int row, int column, double number, in XlsxStyle style);

    public abstract void Bool(int row, int column, bool value, in XlsxStyle style);

    public abstract void Rich(int row, int column, XlsxRichText text);

    public abstract void Formula(int row, int column, string formula, double? result, in XlsxStyle style);
}
