#nullable enable
namespace Divit.Docx;

using Divit.Properties;
using System;
using System.Collections.Generic;

/// <summary>
/// The look of a table, from the document's <see cref="DocxLook"/>.
/// </summary>
public enum DocxTableLook
{
    /// <summary>
    /// The house table: coloured header row, banded rows, light lines.
    /// </summary>
    Banded,

    /// <summary>
    /// Lines around every cell, nothing else: forms and record sheets.
    /// </summary>
    Grid,

    /// <summary>
    /// No lines and no fill: layout of headers, footers and cover pages.
    /// </summary>
    Plain,
}

/// <summary>
/// Where the text of a cell sits between its top and bottom.
/// </summary>
public enum DocxCellAlignment
{
    /// <summary>
    /// At the top.
    /// </summary>
    Top,

    /// <summary>
    /// In the middle.
    /// </summary>
    Center,

    /// <summary>
    /// At the bottom.
    /// </summary>
    Bottom,
}

/// <summary>
/// A table: rows of cells, each cell holding paragraphs. Header rows repeat at the top of every page the table
/// spans; the look colours the first header row and bands the body rows.
/// </summary>
public sealed class DocxTable : DocxBlock
{
    private readonly List<DocxTableRow> _rows = new List<DocxTableRow>();
    private readonly DocxAlignment?[] _alignments;
    private double[]? _widths;

    internal DocxTable(int columns)
    {
        if (columns < 1 || columns > 63)
        {
            throw new ArgumentOutOfRangeException(nameof(columns), Messages.Format(Strings.TableColumnsOutOfRange, columns));
        }

        Columns = columns;
        _alignments = new DocxAlignment?[columns];
    }

    /// <summary>
    /// Number of columns.
    /// </summary>
    public int Columns { get; }

    /// <summary>
    /// The look: banded house table, plain grid, or no lines.
    /// </summary>
    public DocxTableLook Look { get; set; }

    /// <summary>
    /// Centres the table between the margins when it is narrower than the page.
    /// </summary>
    public bool Centered { get; set; }

    /// <summary>
    /// The rows in order.
    /// </summary>
    public IReadOnlyList<DocxTableRow> Rows => _rows;

    /// <summary>
    /// Column widths in centimetres; the width between the margins shared equally when not set.
    /// </summary>
    public IReadOnlyList<double>? ColumnWidths => _widths;

    /// <summary>
    /// Sets the column widths in centimetres, one per column.
    /// </summary>
    /// <param name="widths">The widths.</param>
    /// <returns>This table.</returns>
    /// <exception cref="ArgumentException">The count differs from the number of columns.</exception>
    public DocxTable SetColumnWidths(params double[] widths)
    {
        if (widths.Length != Columns)
        {
            throw new ArgumentException(Messages.Format(Strings.TableWidthsCount, widths.Length, Columns), nameof(widths));
        }

        _widths = widths;
        return this;
    }

    /// <summary>
    /// Aligns the text of a column, for example numbers to the right.
    /// </summary>
    /// <param name="column">Column, 1 for the first.</param>
    /// <param name="alignment">The alignment.</param>
    /// <returns>This table.</returns>
    public DocxTable AlignColumn(int column, DocxAlignment alignment)
    {
        CheckColumn(column);
        _alignments[column - 1] = alignment;
        return this;
    }

    /// <summary>
    /// Adds a header row: bold in the header colours, repeated at the top of each page the table spans.
    /// </summary>
    /// <param name="texts">Text of the cells, one per column; fewer leave the rest empty.</param>
    /// <returns>The row.</returns>
    public DocxTableRow AddHeaderRow(params string?[] texts)
    {
        var row = AddRow(texts);
        row.IsHeader = true;
        return row;
    }

    /// <summary>
    /// Adds a body row.
    /// </summary>
    /// <param name="texts">Text of the cells, one per column; fewer leave the rest empty.</param>
    /// <returns>The row.</returns>
    /// <exception cref="ArgumentException">There are more texts than columns.</exception>
    public DocxTableRow AddRow(params string?[] texts)
    {
        if (texts.Length > Columns)
        {
            throw new ArgumentException(Messages.Format(Strings.TableWidthsCount, texts.Length, Columns), nameof(texts));
        }

        // In a streamed section the rows so far are written now, so a long table is not kept.
        BeforeRow?.Invoke(this);
        var row = new DocxTableRow(this, Columns);
        for (var i = 0; i < texts.Length; i++)
        {
            if (!string.IsNullOrEmpty(texts[i]))
            {
                row.Cells[i].AddParagraph(texts[i]);
            }
        }

        _rows.Add(row);
        return row;
    }

    /// <summary>
    /// Called before a row is added, while the table belongs to a streamed section.
    /// </summary>
    internal Action<DocxTable>? BeforeRow { get; set; }

    /// <summary>
    /// Forgets the rows written so far.
    /// </summary>
    internal void ClearRows() => _rows.Clear();

    internal DocxAlignment? AlignmentOf(int index) => _alignments[index];

    private void CheckColumn(int column)
    {
        if (column < 1 || column > Columns)
        {
            throw new ArgumentOutOfRangeException(nameof(column), Messages.Format(Strings.TableColumnsOutOfRange, column));
        }
    }
}

/// <summary>
/// A row of a table.
/// </summary>
public sealed class DocxTableRow
{
    internal DocxTableRow(DocxTable table, int columns)
    {
        Table = table;
        var cells = new DocxTableCell[columns];
        for (var i = 0; i < columns; i++)
        {
            cells[i] = new DocxTableCell(this, i);
        }

        Cells = cells;
    }

    /// <summary>
    /// The cells, one per column; a cell covered by a neighbour's span is left out when written.
    /// </summary>
    public IReadOnlyList<DocxTableCell> Cells { get; }

    /// <summary>
    /// Whether the row repeats at the top of each page (<see cref="DocxTable.AddHeaderRow"/>).
    /// </summary>
    public bool IsHeader { get; set; }

    /// <summary>
    /// Least height of the row in centimetres; it grows with its content. The content's own height when empty.
    /// </summary>
    public double? Height { get; set; }

    /// <summary>
    /// Keeps the row on one page.
    /// </summary>
    public bool KeepTogether { get; set; } = true;

    internal DocxTable Table { get; }

    /// <summary>
    /// The cell of a column.
    /// </summary>
    /// <param name="column">Column, 1 for the first.</param>
    /// <returns>The cell.</returns>
    public DocxTableCell this[int column] => Cells[column - 1];
}

/// <summary>
/// A cell of a table: paragraphs (and nested tables), fill, alignment, and the columns or rows it spans.
/// </summary>
public sealed class DocxTableCell : DocxContainer
{
    internal DocxTableCell(DocxTableRow row, int index)
    {
        Row = row;
        Index = index;
    }

    /// <summary>
    /// Fill colour; the table look's when empty.
    /// </summary>
    public string? Shading { get; set; }

    /// <summary>
    /// Where the text sits between top and bottom.
    /// </summary>
    public DocxCellAlignment VerticalAlignment { get; set; }

    /// <summary>
    /// Number of columns the cell covers, to the right (1 for itself). The covered cells are left out.
    /// </summary>
    public int ColumnSpan { get; set; } = 1;

    /// <summary>
    /// Number of rows the cell covers, downwards (1 for itself). The cells below are joined to it.
    /// </summary>
    public int RowSpan { get; set; } = 1;

    internal DocxTableRow Row { get; }

    internal int Index { get; }

    /// <summary>
    /// Text of the cell's first paragraph, adding one when the cell is empty; for short cells.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <param name="font">Its font.</param>
    /// <returns>The paragraph.</returns>
    public DocxParagraph SetText(string text, DocxFont? font = null)
    {
        var paragraph = Blocks.Count > 0 && Blocks[0] is DocxParagraph first ? first : AddParagraph();
        return paragraph.Add(text, font);
    }
}
