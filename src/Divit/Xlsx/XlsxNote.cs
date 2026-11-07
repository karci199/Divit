#nullable enable
namespace Divit.Xlsx;

/// <summary>
/// A note on a cell (<see cref="XlsxSheet.Comment"/>): its text and the size of the yellow box it opens in.
/// </summary>
public sealed class XlsxNote
{
    private int _columns = 2;
    private int _rows = 4;

    internal XlsxNote(int row, int column, string text, XlsxRichText? rich, string? author)
    {
        Row = row;
        Column = column;
        Text = text;
        Rich = rich;
        Author = author;
    }

    /// <summary>
    /// Row of the cell (1-based).
    /// </summary>
    public int Row { get; }

    /// <summary>
    /// Column of the cell (1-based).
    /// </summary>
    public int Column { get; }

    /// <summary>
    /// The text, all pieces together.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Author shown in Excel; the workbook author when empty.
    /// </summary>
    public string? Author { get; }

    /// <summary>
    /// Width of the box in columns (1–20); the box opens beside the cell.
    /// </summary>
    public int Columns
    {
        get => _columns;
        set => _columns = System.Math.Max(1, System.Math.Min(20, value));
    }

    /// <summary>
    /// Height of the box in rows (1–100).
    /// </summary>
    public int Rows
    {
        get => _rows;
        set => _rows = System.Math.Max(1, System.Math.Min(100, value));
    }

    /// <summary>
    /// The pieces of a rich note; <c>null</c> for plain text.
    /// </summary>
    internal XlsxRichText? Rich { get; }
}
