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
    /// Underlined text; see <see cref="DoubleUnderline"/> for two lines.
    /// </summary>
    public bool Underline;

    /// <summary>
    /// Underlined with two lines; wins over <see cref="Underline"/>.
    /// </summary>
    public bool DoubleUnderline;

    /// <summary>
    /// Struck-through text.
    /// </summary>
    public bool Strikethrough;

    /// <summary>
    /// Superscript or subscript text; normal when <see cref="XlsxScript.None"/>.
    /// </summary>
    public XlsxScript Script;

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
    /// Shrinks the font until the text fits the cell width; Excel ignores it when <see cref="Wrap"/> is set.
    /// </summary>
    public bool ShrinkToFit;

    /// <summary>
    /// Indent from the aligned edge, in steps of about three spaces (0–250). Takes effect with left, right or
    /// distributed alignment; with no alignment set, left is used.
    /// </summary>
    public int Indent;

    /// <summary>
    /// Text angle in degrees: positive turns counter-clockwise (up to 90), negative clockwise (down to -90). Use
    /// <see cref="VerticalText"/> for letters stacked top to bottom.
    /// </summary>
    public int Rotation;

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
    /// Value of <see cref="Rotation"/> that stacks the letters vertically, one under the other.
    /// </summary>
    public const int VerticalText = 255;

    /// <summary>
    /// Whether no field is set (the workbook's default style, number 0).
    /// </summary>
    public bool IsDefault => Equals(default(XlsxStyle));

    /// <inheritdoc />
    public bool Equals(XlsxStyle other) =>
        FontName == other.FontName && FontSize.Equals(other.FontSize) && Bold == other.Bold && Italic == other.Italic
        && Underline == other.Underline && DoubleUnderline == other.DoubleUnderline && Strikethrough == other.Strikethrough
        && Script == other.Script && FontColor == other.FontColor && Fill == other.Fill
        && Horizontal == other.Horizontal && Vertical == other.Vertical && Wrap == other.Wrap
        && ShrinkToFit == other.ShrinkToFit && Indent == other.Indent && Rotation == other.Rotation
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
            h = h * 397 ^ (Bold ? 1 : 0) ^ (Italic ? 2 : 0) ^ (Underline ? 4 : 0) ^ (Wrap ? 8 : 0)
                ^ (DoubleUnderline ? 16 : 0) ^ (Strikethrough ? 32 : 0) ^ (ShrinkToFit ? 64 : 0) ^ ((int)Script << 7);
            h = h * 397 ^ Indent ^ (Rotation << 8);
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
/// Vertical position of text against the line: normal, raised and small, or lowered and small.
/// </summary>
public enum XlsxScript
{
    /// <summary>
    /// Normal text on the line.
    /// </summary>
    None,

    /// <summary>
    /// Raised small text (x²).
    /// </summary>
    Superscript,

    /// <summary>
    /// Lowered small text (H₂O).
    /// </summary>
    Subscript,
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
