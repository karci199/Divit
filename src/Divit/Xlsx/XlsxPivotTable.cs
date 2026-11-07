#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;

/// <summary>
/// How a pivot table combines the values of a group.
/// </summary>
public enum XlsxPivotFunction
{
    /// <summary>
    /// The sum.
    /// </summary>
    Sum,

    /// <summary>
    /// How many rows the group has.
    /// </summary>
    Count,

    /// <summary>
    /// The average.
    /// </summary>
    Average,

    /// <summary>
    /// The largest value.
    /// </summary>
    Max,

    /// <summary>
    /// The smallest value.
    /// </summary>
    Min,
}

/// <summary>
/// A pivot table (<see cref="XlsxSheet.AddPivotTable"/>): a summary of a data range that the reader rearranges in
/// Excel by dragging fields. Divit writes the data as the pivot cache, the layout, and the summary itself into the
/// cells, so the result shows before Excel refreshes it; Excel refreshes it when the file opens.
/// </summary>
/// <remarks>
/// The first row of the source range names the fields; the rows below are the records. Give the fields by these
/// names: one on the rows, optionally one on the columns and one as a filter, and one whose values are combined.
/// </remarks>
public sealed class XlsxPivotTable
{
    internal XlsxPivotTable(XlsxSheet sheet, int row, int column, XlsxSheet source, int row1, int column1, int row2, int column2, string name)
    {
        Sheet = sheet;
        Row = row;
        Column = column;
        Source = source;
        SourceRow1 = row1;
        SourceColumn1 = column1;
        SourceRow2 = row2;
        SourceColumn2 = column2;
        Name = name;
        Fields = Array.Empty<Field>();
    }

    /// <summary>
    /// Name of the pivot table, unique on its sheet.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// The field whose values become the rows; required.
    /// </summary>
    public string? RowField { get; set; }

    /// <summary>
    /// The field whose values become the columns; none when empty (one column of values).
    /// </summary>
    public string? ColumnField { get; set; }

    /// <summary>
    /// A field offered as a filter above the table (all values shown); none when empty.
    /// </summary>
    public string? FilterField { get; set; }

    /// <summary>
    /// The field whose values are combined; required.
    /// </summary>
    public string? ValueField { get; set; }

    /// <summary>
    /// How the values are combined.
    /// </summary>
    public XlsxPivotFunction Function { get; set; }

    /// <summary>
    /// Number format of the values, for example <c>#,##0.00</c>; General when empty.
    /// </summary>
    public string? NumberFormat { get; set; }

    /// <summary>
    /// The pivot table style, one of Excel's <c>PivotStyleLight1</c>–<c>28</c>, <c>PivotStyleMedium1</c>–<c>28</c>,
    /// <c>PivotStyleDark1</c>–<c>28</c>.
    /// </summary>
    public string Style { get; set; } = "PivotStyleLight16";

    internal XlsxSheet Sheet { get; }

    internal int Row { get; }

    internal int Column { get; }

    internal XlsxSheet Source { get; }

    internal int SourceRow1 { get; }

    internal int SourceColumn1 { get; }

    internal int SourceRow2 { get; }

    internal int SourceColumn2 { get; }

    /// <summary>
    /// Number of the cache and of the table in the workbook, from 1; set when saving.
    /// </summary>
    internal int Id { get; set; }

    internal Field[] Fields { get; private set; }

    internal int RecordCount => SourceRow2 - SourceRow1;

    internal int RowIndex { get; private set; }

    internal int ColumnIndex { get; private set; } = -1;

    internal int FilterIndex { get; private set; } = -1;

    internal int ValueIndex { get; private set; }

    /// <summary>
    /// Items of the row field in display order, as indexes into its shared items.
    /// </summary>
    internal int[] RowOrder { get; private set; } = Array.Empty<int>();

    internal int[] ColumnOrder { get; private set; } = Array.Empty<int>();

    /// <summary>
    /// Top-left cell of the table body (below the filter, when there is one).
    /// </summary>
    internal int BodyRow => FilterIndex >= 0 ? Row + 2 : Row;

    /// <summary>
    /// Last row and column of the table body.
    /// </summary>
    internal int LastRow => BodyRow + (ColumnIndex >= 0 ? 2 : 1) + RowOrder.Length;

    internal int LastColumn => Column + (ColumnIndex >= 0 ? ColumnOrder.Length + 1 : 1);

    /// <summary>
    /// The caption of the values, as Excel writes it: "Sum of Amount".
    /// </summary>
    internal string ValueCaption
    {
        get
        {
            var name = Fields[ValueIndex].Name;
            switch (Function)
            {
                case XlsxPivotFunction.Count: return "Count of " + name;
                case XlsxPivotFunction.Average: return "Average of " + name;
                case XlsxPivotFunction.Max: return "Max of " + name;
                case XlsxPivotFunction.Min: return "Min of " + name;
                default: return "Sum of " + name;
            }
        }
    }

    /// <summary>
    /// Reads the source into the cache model, checks the fields and writes the summary into the target cells.
    /// </summary>
    /// <exception cref="InvalidOperationException">A field is missing or unknown, used twice, or the value field is
    /// not numeric for a function other than count.</exception>
    internal void Prepare()
    {
        var count = SourceColumn2 - SourceColumn1 + 1;
        var fields = new Field[count];
        for (var i = 0; i < count; i++)
        {
            var column = SourceColumn1 + i;
            var name = Source.TextAt(SourceRow1, column);
            fields[i] = new Field(string.IsNullOrEmpty(name) ? "Column" + (i + 1).ToString(CultureInfo.InvariantCulture) : name!);
            for (var r = SourceRow1 + 1; r <= SourceRow2; r++)
            {
                fields[i].Add(Source.NumberAt(r, column), Source.NumberAt(r, column) == null ? Source.TextAt(r, column) : null);
            }
        }

        Fields = fields;
        RowIndex = Find(RowField, true);
        ColumnIndex = Find(ColumnField, false);
        FilterIndex = Find(FilterField, false);
        ValueIndex = Find(ValueField, true);
        var used = new HashSet<int>();
        foreach (var index in new[] { RowIndex, ColumnIndex, FilterIndex, ValueIndex })
        {
            if (index >= 0 && !used.Add(index))
            {
                throw new InvalidOperationException(Messages.Format(Strings.PivotFieldTwice, fields[index].Name, Name));
            }
        }

        if (Function != XlsxPivotFunction.Count && !fields[ValueIndex].Numeric)
        {
            throw new InvalidOperationException(Messages.Format(Strings.PivotValueNotNumeric, fields[ValueIndex].Name, Name));
        }

        // Axis fields list their items; the value field keeps its numbers in the records.
        foreach (var index in new[] { RowIndex, ColumnIndex, FilterIndex })
        {
            if (index >= 0)
            {
                fields[index].OnAxis = true;
            }
        }

        RowOrder = fields[RowIndex].SortedItems();
        ColumnOrder = ColumnIndex >= 0 ? fields[ColumnIndex].SortedItems() : Array.Empty<int>();
        WriteCells();
    }

    /// <summary>
    /// The result of one group: the records whose row item (and column item) match; -1 matches all.
    /// </summary>
    internal double? Result(int rowItem, int columnItem)
    {
        var values = Fields[ValueIndex];
        var n = 0;
        double sum = 0, min = double.MaxValue, max = double.MinValue;
        for (var r = 0; r < RecordCount; r++)
        {
            if ((rowItem >= 0 && Fields[RowIndex].Records[r] != rowItem)
                || (columnItem >= 0 && ColumnIndex >= 0 && Fields[ColumnIndex].Records[r] != columnItem))
            {
                continue;
            }

            if (Function == XlsxPivotFunction.Count)
            {
                if (values.HasValue(r)) { n++; }
                continue;
            }

            if (values.Numbers[r] is double v)
            {
                n++;
                sum += v;
                min = Math.Min(min, v);
                max = Math.Max(max, v);
            }
        }

        switch (Function)
        {
            case XlsxPivotFunction.Count: return n == 0 ? (double?)null : n;
            case XlsxPivotFunction.Average: return n == 0 ? (double?)null : sum / n;
            case XlsxPivotFunction.Max: return n == 0 ? (double?)null : max;
            case XlsxPivotFunction.Min: return n == 0 ? (double?)null : min;
            default: return n == 0 ? (double?)null : sum;
        }
    }

    /// <summary>
    /// The summary in the cells, in Excel's compact layout: caption and "Column Labels", then "Row Labels" with the
    /// column items, the rows, and the grand totals.
    /// </summary>
    private void WriteCells()
    {
        var bold = new XlsxStyle { Bold = true };
        var values = new XlsxStyle { NumberFormat = NumberFormat };
        if (FilterIndex >= 0)
        {
            Sheet.Set(Row, Column, Fields[FilterIndex].Name);
            Sheet.Set(Row, Column + 1, "(All)");
        }

        var row = BodyRow;
        var columns = ColumnIndex >= 0 ? ColumnOrder : new[] { -1 };
        if (ColumnIndex >= 0)
        {
            Sheet.Set(row, Column, ValueCaption);
            Sheet.Set(row, Column + 1, "Column Labels");
            row++;
            Sheet.Set(row, Column, "Row Labels");
            for (var c = 0; c < ColumnOrder.Length; c++)
            {
                Label(row, Column + 1 + c, Fields[ColumnIndex], ColumnOrder[c]);
            }

            Sheet.Set(row, Column + 1 + ColumnOrder.Length, "Grand Total");
        }
        else
        {
            Sheet.Set(row, Column, "Row Labels");
            Sheet.Set(row, Column + 1, ValueCaption);
        }

        Sheet.SetStyle(row, Column, row, LastColumn, bold);
        foreach (var item in RowOrder)
        {
            row++;
            Label(row, Column, Fields[RowIndex], item);
            for (var c = 0; c < columns.Length; c++)
            {
                Value(row, Column + 1 + c, Result(item, columns[c]), values);
            }

            if (ColumnIndex >= 0)
            {
                Value(row, Column + 1 + columns.Length, Result(item, -1), values);
            }
        }

        row++;
        Sheet.Set(row, Column, "Grand Total");
        for (var c = 0; c < columns.Length; c++)
        {
            Value(row, Column + 1 + c, Result(-1, columns[c]), values);
        }

        if (ColumnIndex >= 0)
        {
            Value(row, Column + 1 + columns.Length, Result(-1, -1), values);
        }

        Sheet.Format(row, Column, row, LastColumn, (ref XlsxStyle s) => s.Bold = true);
    }

    /// <summary>
    /// An item label: a number stays a number in its cell, as Excel writes it.
    /// </summary>
    private void Label(int row, int column, Field field, int item)
    {
        if (field.Items[item] is double number)
        {
            Sheet.Set(row, column, number);
        }
        else
        {
            Sheet.Set(row, column, field.Label(item));
        }
    }

    private void Value(int row, int column, double? value, XlsxStyle style)
    {
        if (value is double v)
        {
            Sheet.Set(row, column, v);
        }

        if (style.NumberFormat != null)
        {
            Sheet.Format(row, column, (ref XlsxStyle s) => s.NumberFormat = style.NumberFormat);
        }
    }

    private int Find(string? name, bool required)
    {
        if (name == null)
        {
            if (required)
            {
                throw new InvalidOperationException(Messages.Format(Strings.PivotFieldMissing, Name));
            }

            return -1;
        }

        for (var i = 0; i < Fields.Length; i++)
        {
            if (string.Equals(Fields[i].Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        throw new InvalidOperationException(Messages.Format(Strings.PivotFieldUnknown, name, Name));
    }

    /// <summary>
    /// One column of the source in the cache: its distinct items in order of first appearance, and per record the
    /// item it holds (or its number).
    /// </summary>
    internal sealed class Field
    {
        private readonly Dictionary<string, int> _texts = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly Dictionary<double, int> _numbers = new Dictionary<double, int>();
        private int _blank = -1;

        public Field(string name)
        {
            Name = name;
        }

        public string Name { get; }

        /// <summary>
        /// The distinct values: a string, a double, or <c>null</c> for an empty cell.
        /// </summary>
        public List<object?> Items { get; } = new List<object?>();

        /// <summary>
        /// For each record, the index of its item.
        /// </summary>
        public List<int> Records { get; } = new List<int>();

        /// <summary>
        /// For each record, its number; <c>null</c> for text or empty.
        /// </summary>
        public List<double?> Numbers { get; } = new List<double?>();

        public bool OnAxis { get; set; }

        public bool HasText { get; private set; }

        public bool HasNumber { get; private set; }

        public bool HasBlank => _blank >= 0;

        /// <summary>
        /// Whether every value is a number (empty cells allowed).
        /// </summary>
        public bool Numeric => HasNumber && !HasText;

        public bool HasValue(int record) => Items[Records[record]] != null;

        public void Add(double? number, string? text)
        {
            Numbers.Add(number);
            int index;
            if (number is double n)
            {
                HasNumber = true;
                if (!_numbers.TryGetValue(n, out index))
                {
                    index = Items.Count;
                    Items.Add(n);
                    _numbers[n] = index;
                }
            }
            else if (!string.IsNullOrEmpty(text))
            {
                HasText = true;
                if (!_texts.TryGetValue(text!, out index))
                {
                    index = Items.Count;
                    Items.Add(text);
                    _texts[text!] = index;
                }
            }
            else
            {
                if (_blank < 0)
                {
                    _blank = Items.Count;
                    Items.Add(null);
                }

                index = _blank;
            }

            Records.Add(index);
        }

        /// <summary>
        /// The items in display order: numbers ascending, then texts by the invariant culture, then the empty item.
        /// </summary>
        public int[] SortedItems()
        {
            var order = new List<int>();
            for (var i = 0; i < Items.Count; i++)
            {
                order.Add(i);
            }

            order.Sort((a, b) =>
            {
                var x = Items[a];
                var y = Items[b];
                if (x == null || y == null)
                {
                    return x == null ? (y == null ? 0 : 1) : -1;
                }

                if (x is double dx && y is double dy)
                {
                    return dx.CompareTo(dy);
                }

                if (x is double)
                {
                    return -1;
                }

                if (y is double)
                {
                    return 1;
                }

                var compare = string.Compare((string)x, (string)y, StringComparison.InvariantCultureIgnoreCase);
                return compare != 0 ? compare : string.CompareOrdinal((string)x, (string)y);
            });
            return order.ToArray();
        }

        /// <summary>
        /// The label of an item as a cell shows it.
        /// </summary>
        public string Label(int item) =>
            Items[item] is double d ? d.ToString("R", CultureInfo.InvariantCulture) : (string?)Items[item] ?? "(blank)";
    }
}
