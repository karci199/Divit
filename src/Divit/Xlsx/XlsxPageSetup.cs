#nullable enable
namespace Divit.Xlsx;

/// <summary>
/// Whether a sheet tab is shown.
/// </summary>
public enum XlsxSheetVisibility
{
    /// <summary>
    /// The tab is shown.
    /// </summary>
    Visible,

    /// <summary>
    /// The tab is hidden; the user can show it again with Unhide.
    /// </summary>
    Hidden,

    /// <summary>
    /// The tab is hidden and missing from Unhide; only the VBA editor or a program can show it again (for example a
    /// sheet of lookup data).
    /// </summary>
    VeryHidden,
}

/// <summary>
/// Paper sizes; the value is Excel's paper number.
/// </summary>
public enum XlsxPaperSize
{
    /// <summary>
    /// US Letter, 8.5 × 11 in.
    /// </summary>
    Letter = 1,

    /// <summary>
    /// US Legal, 8.5 × 14 in.
    /// </summary>
    Legal = 5,

    /// <summary>
    /// A3, 297 × 420 mm.
    /// </summary>
    A3 = 8,

    /// <summary>
    /// A4, 210 × 297 mm.
    /// </summary>
    A4 = 9,

    /// <summary>
    /// A5, 148 × 210 mm.
    /// </summary>
    A5 = 11,

    /// <summary>
    /// B5 (JIS), 182 × 257 mm.
    /// </summary>
    B5 = 13,
}

/// <summary>
/// Page margins of a printed sheet, in inches (2.54 cm).
/// </summary>
public readonly struct XlsxMargins
{
    /// <summary>
    /// Creates margins.
    /// </summary>
    /// <param name="left">Left margin.</param>
    /// <param name="right">Right margin.</param>
    /// <param name="top">Top margin.</param>
    /// <param name="bottom">Bottom margin.</param>
    /// <param name="header">Distance of the header from the top edge.</param>
    /// <param name="footer">Distance of the footer from the bottom edge.</param>
    public XlsxMargins(double left, double right, double top, double bottom, double header, double footer)
    {
        Left = left;
        Right = right;
        Top = top;
        Bottom = bottom;
        Header = header;
        Footer = footer;
    }

    /// <summary>
    /// Divit's default: half an inch around, headers 0.3 in from the edge.
    /// </summary>
    public static XlsxMargins Default => new XlsxMargins(0.5, 0.5, 0.5, 0.5, 0.3, 0.3);

    /// <summary>
    /// Excel's "Normal" margins.
    /// </summary>
    public static XlsxMargins Normal => new XlsxMargins(0.7, 0.7, 0.75, 0.75, 0.3, 0.3);

    /// <summary>
    /// Excel's "Narrow" margins.
    /// </summary>
    public static XlsxMargins Narrow => new XlsxMargins(0.25, 0.25, 0.75, 0.75, 0.3, 0.3);

    /// <summary>
    /// Excel's "Wide" margins.
    /// </summary>
    public static XlsxMargins Wide => new XlsxMargins(1, 1, 1, 1, 0.5, 0.5);

    /// <summary>
    /// Left margin in inches.
    /// </summary>
    public double Left { get; }

    /// <summary>
    /// Right margin in inches.
    /// </summary>
    public double Right { get; }

    /// <summary>
    /// Top margin in inches.
    /// </summary>
    public double Top { get; }

    /// <summary>
    /// Bottom margin in inches.
    /// </summary>
    public double Bottom { get; }

    /// <summary>
    /// Distance of the header from the top edge, in inches.
    /// </summary>
    public double Header { get; }

    /// <summary>
    /// Distance of the footer from the bottom edge, in inches.
    /// </summary>
    public double Footer { get; }

    /// <summary>
    /// Margins given in centimetres.
    /// </summary>
    /// <param name="left">Left margin.</param>
    /// <param name="right">Right margin.</param>
    /// <param name="top">Top margin.</param>
    /// <param name="bottom">Bottom margin.</param>
    /// <param name="header">Distance of the header from the top edge.</param>
    /// <param name="footer">Distance of the footer from the bottom edge.</param>
    /// <returns>The margins in inches.</returns>
    public static XlsxMargins FromCentimeters(double left, double right, double top, double bottom, double header, double footer) =>
        new XlsxMargins(left / 2.54, right / 2.54, top / 2.54, bottom / 2.54, header / 2.54, footer / 2.54);
}

/// <summary>
/// What users may still do on a locked sheet (<see cref="XlsxSheet.Protect"/>); selecting cells is always allowed.
/// </summary>
[System.Flags]
public enum XlsxAllow
{
    /// <summary>
    /// Nothing beyond selecting and editing unlocked cells.
    /// </summary>
    None = 0,

    /// <summary>
    /// Changing the format of cells.
    /// </summary>
    FormatCells = 1,

    /// <summary>
    /// Changing column widths, hiding columns.
    /// </summary>
    FormatColumns = 2,

    /// <summary>
    /// Changing row heights, hiding rows.
    /// </summary>
    FormatRows = 4,

    /// <summary>
    /// Inserting columns.
    /// </summary>
    InsertColumns = 8,

    /// <summary>
    /// Inserting rows.
    /// </summary>
    InsertRows = 16,

    /// <summary>
    /// Inserting links.
    /// </summary>
    InsertHyperlinks = 32,

    /// <summary>
    /// Deleting columns.
    /// </summary>
    DeleteColumns = 64,

    /// <summary>
    /// Deleting rows.
    /// </summary>
    DeleteRows = 128,

    /// <summary>
    /// Sorting (only ranges of unlocked cells).
    /// </summary>
    Sort = 256,

    /// <summary>
    /// Using the drop-downs of an existing auto filter.
    /// </summary>
    AutoFilter = 512,

    /// <summary>
    /// Using pivot tables.
    /// </summary>
    PivotTables = 1024,
}
