#nullable enable
namespace Divit.Xlsx;

using System;

/// <summary>
/// The look of a cell: font, fill, alignment, number format and borders.
/// </summary>
/// <remarks>
/// A value type: cells that look the same share one style record when saved (the style table of
/// <see cref="XlsxWorkbook"/>). An empty field means "default": no font name uses the workbook font, no colour is black,
/// no fill means none. Colours are <c>#RRGGBB</c>.
/// </remarks>
public struct XlsxStyle : IEquatable<XlsxStyle>
{
    /// <summary>
    /// Font name; the workbook default when empty.
    /// </summary>
    public string? FontName;

    /// <summary>
    /// Font size in points; 0 uses the workbook default.
    /// </summary>
    public double FontSize;

    /// <summary>
    /// Bold text.
    /// </summary>
    public bool Bold;

    /// <summary>
    /// Italic text.
    /// </summary>
    public bool Italic;

    /// <summary>
    /// Underlined text.
    /// </summary>
    public bool Underline;

    /// <summary>
    /// Text colour as <c>#RRGGBB</c>; black when empty.
    /// </summary>
    public string? FontColor;

    /// <summary>
    /// Fill colour as <c>#RRGGBB</c>; no fill when empty.
    /// </summary>
    public string? Fill;

    /// <summary>
    /// Horizontal alignment: <c>left</c>, <c>center</c>, <c>right</c>; general when empty (text left, numbers right).
    /// </summary>
    public string? Horizontal;

    /// <summary>
    /// Vertical alignment: <c>top</c>, <c>center</c>, <c>bottom</c>; bottom when empty.
    /// </summary>
    public string? Vertical;

    /// <summary>
    /// Wraps text onto further lines when it does not fit the cell.
    /// </summary>
    public bool Wrap;

    /// <summary>
    /// Number format as an Excel code (for example <c>0.00</c>); General when empty.
    /// </summary>
    public string? NumberFormat;

    /// <summary>
    /// Left border.
    /// </summary>
    public XlsxBorder Left;

    /// <summary>
    /// Right border.
    /// </summary>
    public XlsxBorder Right;

    /// <summary>
    /// Top border.
    /// </summary>
    public XlsxBorder Top;

    /// <summary>
    /// Bottom border.
    /// </summary>
    public XlsxBorder Bottom;

    /// <summary>
    /// Whether no field is set (the workbook's default style, number 0).
    /// </summary>
    public bool IsDefault => Equals(default(XlsxStyle));

    /// <inheritdoc />
    public bool Equals(XlsxStyle other) =>
        FontName == other.FontName && FontSize.Equals(other.FontSize) && Bold == other.Bold && Italic == other.Italic
        && Underline == other.Underline && FontColor == other.FontColor && Fill == other.Fill
        && Horizontal == other.Horizontal && Vertical == other.Vertical && Wrap == other.Wrap
        && NumberFormat == other.NumberFormat && Left.Equals(other.Left) && Right.Equals(other.Right)
        && Top.Equals(other.Top) && Bottom.Equals(other.Bottom);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is XlsxStyle other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            var h = (FontName?.GetHashCode() ?? 0) * 397 ^ FontSize.GetHashCode();
            h = h * 397 ^ (Bold ? 1 : 0) ^ (Italic ? 2 : 0) ^ (Underline ? 4 : 0) ^ (Wrap ? 8 : 0);
            h = h * 397 ^ (FontColor?.GetHashCode() ?? 0);
            h = h * 397 ^ (Fill?.GetHashCode() ?? 0);
            h = h * 397 ^ (Horizontal?.GetHashCode() ?? 0);
            h = h * 397 ^ (Vertical?.GetHashCode() ?? 0);
            h = h * 397 ^ (NumberFormat?.GetHashCode() ?? 0);
            h = h * 397 ^ Left.GetHashCode() ^ Right.GetHashCode() * 3 ^ Top.GetHashCode() * 7 ^ Bottom.GetHashCode() * 11;
            return h;
        }
    }
}

/// <summary>
/// One edge of a cell: line kind and colour. The default has no line.
/// </summary>
public readonly struct XlsxBorder : IEquatable<XlsxBorder>
{
    /// <summary>
    /// Creates a border.
    /// </summary>
    /// <param name="line">Excel line kind: <c>thin</c>, <c>medium</c>, <c>thick</c>, <c>dashed</c> …</param>
    /// <param name="color">Colour as <c>#RRGGBB</c>; black when empty.</param>
    public XlsxBorder(string line, string? color)
    {
        Line = line;
        Color = color;
    }

    /// <summary>
    /// Line kind; no border when empty.
    /// </summary>
    public string? Line { get; }

    /// <summary>
    /// Line colour as <c>#RRGGBB</c>; black when empty.
    /// </summary>
    public string? Color { get; }

    /// <inheritdoc />
    public bool Equals(XlsxBorder other) => Line == other.Line && Color == other.Color;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is XlsxBorder other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => ((Line?.GetHashCode() ?? 0) * 397) ^ (Color?.GetHashCode() ?? 0);
}

/// <summary>
/// Changes the style of a cell in place; edits stack up (fill first, then borders …).
/// </summary>
/// <param name="style">The cell's current style.</param>
public delegate void XlsxStyleEdit(ref XlsxStyle style);
