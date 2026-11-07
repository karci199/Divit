#nullable enable
namespace Divit.Docx;

/// <summary>
/// Paper sizes.
/// </summary>
public enum DocxPageSize
{
    /// <summary>
    /// A4, 210 × 297 mm.
    /// </summary>
    A4,

    /// <summary>
    /// A3, 297 × 420 mm.
    /// </summary>
    A3,

    /// <summary>
    /// A5, 148 × 210 mm.
    /// </summary>
    A5,

    /// <summary>
    /// US Letter, 8.5 × 11 in.
    /// </summary>
    Letter,

    /// <summary>
    /// US Legal, 8.5 × 14 in.
    /// </summary>
    Legal,
}

/// <summary>
/// How page numbers are written.
/// </summary>
public enum DocxNumberFormat
{
    /// <summary>
    /// 1, 2, 3.
    /// </summary>
    Decimal,

    /// <summary>
    /// i, ii, iii, often for the front pages.
    /// </summary>
    LowerRoman,

    /// <summary>
    /// I, II, III.
    /// </summary>
    UpperRoman,
}

/// <summary>
/// Where the text of a page sits between the top and bottom margins.
/// </summary>
public enum DocxVerticalAlignment
{
    /// <summary>
    /// From the top.
    /// </summary>
    Top,

    /// <summary>
    /// In the middle, as on a cover page.
    /// </summary>
    Center,

    /// <summary>
    /// At the bottom.
    /// </summary>
    Bottom,
}

/// <summary>
/// Page margins in centimetres, and the distances of header and footer from the paper's edge.
/// </summary>
public readonly struct DocxMargins
{
    /// <summary>
    /// Creates margins.
    /// </summary>
    /// <param name="top">Top margin.</param>
    /// <param name="bottom">Bottom margin.</param>
    /// <param name="left">Left margin.</param>
    /// <param name="right">Right margin.</param>
    /// <param name="header">Distance of the header from the top edge.</param>
    /// <param name="footer">Distance of the footer from the bottom edge.</param>
    public DocxMargins(double top, double bottom, double left, double right, double header = 1.25, double footer = 1.25)
    {
        Top = top;
        Bottom = bottom;
        Left = left;
        Right = right;
        Header = header;
        Footer = footer;
    }

    /// <summary>
    /// 2.5 cm around, header and footer 1.25 cm from the edge.
    /// </summary>
    public static DocxMargins Normal => new DocxMargins(2.5, 2.5, 2.5, 2.5);

    /// <summary>
    /// Top margin in centimetres.
    /// </summary>
    public double Top { get; }

    /// <summary>
    /// Bottom margin in centimetres.
    /// </summary>
    public double Bottom { get; }

    /// <summary>
    /// Left margin in centimetres.
    /// </summary>
    public double Left { get; }

    /// <summary>
    /// Right margin in centimetres.
    /// </summary>
    public double Right { get; }

    /// <summary>
    /// Distance of the header from the top edge in centimetres.
    /// </summary>
    public double Header { get; }

    /// <summary>
    /// Distance of the footer from the bottom edge in centimetres.
    /// </summary>
    public double Footer { get; }
}

/// <summary>
/// A header or footer: paragraphs (and tables) repeated on the pages of a section. Fields such as
/// <see cref="DocxField.Page"/> show the page they are on.
/// </summary>
public sealed class DocxHeaderFooter : DocxContainer
{
    internal DocxHeaderFooter(bool footer)
    {
        IsFooter = footer;
    }

    internal bool IsFooter { get; }

    /// <summary>
    /// Number of the part in the package, set when saving.
    /// </summary>
    internal int Number { get; set; }

    internal override string DefaultStyle => DocxStyle.NoSpacing;
}

/// <summary>
/// A section of a document: pages with the same size, margins, headers and footers. A new section starts on a new
/// page (<see cref="DocxDocument.AddSection"/>); a cover page is often a section of its own.
/// </summary>
/// <remarks>
/// A section without its own header or footer continues the previous section's, as Word's "Link to Previous" does.
/// </remarks>
public sealed class DocxSection : DocxContainer
{
    private DocxHeaderFooter? _header;
    private DocxHeaderFooter? _footer;
    private DocxHeaderFooter? _firstHeader;
    private DocxHeaderFooter? _firstFooter;
    private DocxHeaderFooter? _evenHeader;
    private DocxHeaderFooter? _evenFooter;

    internal DocxSection(DocxSection? previous)
    {
        if (previous != null)
        {
            PageSize = previous.PageSize;
            Landscape = previous.Landscape;
            Margins = previous.Margins;
        }
    }

    /// <summary>
    /// Paper size.
    /// </summary>
    public DocxPageSize PageSize { get; set; } = DocxPageSize.A4;

    /// <summary>
    /// Turns the page sideways (for a wide table).
    /// </summary>
    public bool Landscape { get; set; }

    /// <summary>
    /// Page margins.
    /// </summary>
    public DocxMargins Margins { get; set; } = DocxMargins.Normal;

    /// <summary>
    /// Where text sits between top and bottom margin.
    /// </summary>
    public DocxVerticalAlignment VerticalAlignment { get; set; }

    /// <summary>
    /// Number of the section's first page; it continues from the previous section when empty.
    /// </summary>
    public int? PageNumberStart { get; set; }

    /// <summary>
    /// How the section's page numbers are written.
    /// </summary>
    public DocxNumberFormat PageNumberFormat { get; set; }

    /// <summary>
    /// Text set large and slanted behind the pages, such as "DRAFT" or "CONFIDENTIAL"; none when empty. It is part of
    /// the section's headers, so the section gets a header of its own when it has none.
    /// </summary>
    public string? Watermark { get; set; }

    /// <summary>
    /// Colour of the watermark, drawn half transparent.
    /// </summary>
    public string WatermarkColor { get; set; } = "#C0C0C0";

    /// <summary>
    /// The header of every page (or the odd pages when <see cref="EvenPageHeader"/> is used). Created when first used.
    /// </summary>
    public DocxHeaderFooter Header => _header ??= new DocxHeaderFooter(false);

    /// <summary>
    /// The footer of every page (or the odd pages). Created when first used.
    /// </summary>
    public DocxHeaderFooter Footer => _footer ??= new DocxHeaderFooter(true);

    /// <summary>
    /// The header of the section's first page only; using it gives the first page its own header and footer (left
    /// empty, they are blank, as on a cover page).
    /// </summary>
    public DocxHeaderFooter FirstPageHeader => _firstHeader ??= new DocxHeaderFooter(false);

    /// <summary>
    /// The footer of the section's first page only, for example a letterhead under a cover page.
    /// </summary>
    public DocxHeaderFooter FirstPageFooter => _firstFooter ??= new DocxHeaderFooter(true);

    /// <summary>
    /// The header of the even pages, for double-sided printing.
    /// </summary>
    public DocxHeaderFooter EvenPageHeader => _evenHeader ??= new DocxHeaderFooter(false);

    /// <summary>
    /// The footer of the even pages.
    /// </summary>
    public DocxHeaderFooter EvenPageFooter => _evenFooter ??= new DocxHeaderFooter(true);

    internal DocxHeaderFooter? DefaultHeaderPart => _header;

    internal DocxHeaderFooter? DefaultFooterPart => _footer;

    internal DocxHeaderFooter? FirstHeaderPart => _firstHeader;

    internal DocxHeaderFooter? FirstFooterPart => _firstFooter;

    internal DocxHeaderFooter? EvenHeaderPart => _evenHeader;

    internal DocxHeaderFooter? EvenFooterPart => _evenFooter;

    internal bool DifferentFirstPage => _firstHeader != null || _firstFooter != null;

    /// <summary>
    /// Paper width and height in twentieths of a point, turned for landscape.
    /// </summary>
    internal (int Width, int Height) PaperTwips
    {
        get
        {
            var (w, h) = PageSize switch
            {
                DocxPageSize.A3 => (16838, 23811),
                DocxPageSize.A5 => (8391, 11906),
                DocxPageSize.Letter => (12240, 15840),
                DocxPageSize.Legal => (12240, 20160),
                _ => (11906, 16838),
            };
            return Landscape ? (h, w) : (w, h);
        }
    }

    /// <summary>
    /// Width between the margins in twentieths of a point: where right-aligned tab stops go.
    /// </summary>
    internal int TextWidthTwips => PaperTwips.Width - DocxUnits.Twips(Margins.Left) - DocxUnits.Twips(Margins.Right);
}

/// <summary>
/// Conversions to WordprocessingML's units.
/// </summary>
internal static class DocxUnits
{
    /// <summary>
    /// Centimetres to twentieths of a point (1 cm = 567 twips).
    /// </summary>
    public static int Twips(double centimetres) => (int)System.Math.Round(centimetres * 1440 / 2.54);

    /// <summary>
    /// Points to twentieths of a point.
    /// </summary>
    public static int PointTwips(double points) => (int)System.Math.Round(points * 20);

    /// <summary>
    /// Points to half-points, the unit of font sizes.
    /// </summary>
    public static int HalfPoints(double points) => (int)System.Math.Round(points * 2);

    /// <summary>
    /// Points to eighths of a point, the unit of border widths (2–96).
    /// </summary>
    public static int Eighths(double points) => System.Math.Max(2, System.Math.Min(96, (int)System.Math.Round(points * 8)));
}
