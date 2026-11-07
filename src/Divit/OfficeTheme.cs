#nullable enable
namespace Divit;

/// <summary>
/// The colour palette and fonts of a document's theme, as Excel, Word and PowerPoint show them in their colour and
/// font pickers. Colours are <c>#RRGGBB</c>.
/// </summary>
/// <remarks>
/// Styles in Divit name colours directly; use these properties to keep the content in the theme's palette:
/// <c>new XlsxStyle { Fill = book.Theme.Accent1 }</c>.
/// </remarks>
public sealed class OfficeTheme
{
    /// <summary>
    /// Name shown in the Office theme gallery.
    /// </summary>
    public string Name { get; set; } = "Divit";

    /// <summary>
    /// Text on light backgrounds.
    /// </summary>
    public string Dark1 { get; set; } = "#000000";

    /// <summary>
    /// Background.
    /// </summary>
    public string Light1 { get; set; } = "#FFFFFF";

    /// <summary>
    /// Second dark colour (headings).
    /// </summary>
    public string Dark2 { get; set; } = "#44546A";

    /// <summary>
    /// Second light colour.
    /// </summary>
    public string Light2 { get; set; } = "#E7E6E6";

    /// <summary>
    /// First accent colour.
    /// </summary>
    public string Accent1 { get; set; } = "#4472C4";

    /// <summary>
    /// Second accent colour.
    /// </summary>
    public string Accent2 { get; set; } = "#ED7D31";

    /// <summary>
    /// Third accent colour.
    /// </summary>
    public string Accent3 { get; set; } = "#A5A5A5";

    /// <summary>
    /// Fourth accent colour.
    /// </summary>
    public string Accent4 { get; set; } = "#FFC000";

    /// <summary>
    /// Fifth accent colour.
    /// </summary>
    public string Accent5 { get; set; } = "#5B9BD5";

    /// <summary>
    /// Sixth accent colour.
    /// </summary>
    public string Accent6 { get; set; } = "#70AD47";

    /// <summary>
    /// Links.
    /// </summary>
    public string Hyperlink { get; set; } = "#0563C1";

    /// <summary>
    /// Links already followed.
    /// </summary>
    public string FollowedHyperlink { get; set; } = "#954F72";

    /// <summary>
    /// Font of headings; the document's default font when empty.
    /// </summary>
    public string? HeadingFont { get; set; }

    /// <summary>
    /// Font of body text; the document's default font when empty.
    /// </summary>
    public string? BodyFont { get; set; }
}
