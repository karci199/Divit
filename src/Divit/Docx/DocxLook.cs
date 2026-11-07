#nullable enable
namespace Divit.Docx;

using System.Collections.Generic;

/// <summary>
/// The house style of a document: fonts, sizes and colours of body text, headings, captions and tables. The styles
/// Word shows (Normal, Heading 1 …) are built from it when the document is saved; change it once and every
/// paragraph follows.
/// </summary>
public sealed class DocxLook
{
    /// <summary>
    /// Font of body text.
    /// </summary>
    public string BodyFont { get; set; } = "Calibri";

    /// <summary>
    /// Size of body text in points.
    /// </summary>
    public double BodySize { get; set; } = 11;

    /// <summary>
    /// Colour of body text.
    /// </summary>
    public string BodyColor { get; set; } = "#262626";

    /// <summary>
    /// Font of headings and the title; the body font when empty.
    /// </summary>
    public string? HeadingFont { get; set; }

    /// <summary>
    /// Colour of headings and the title.
    /// </summary>
    public string HeadingColor { get; set; } = "#1F4E79";

    /// <summary>
    /// Sizes of heading levels 1–6 in points; levels past the list take its last size.
    /// </summary>
    public IReadOnlyList<double> HeadingSizes { get; set; } = new[] { 16.0, 14, 12, 11, 11, 11 };

    /// <summary>
    /// Size of the title in points.
    /// </summary>
    public double TitleSize { get; set; } = 26;

    /// <summary>
    /// Space after a body paragraph in points.
    /// </summary>
    public double SpaceAfter { get; set; } = 6;

    /// <summary>
    /// Line spacing of body text as a multiple of single spacing.
    /// </summary>
    public double LineSpacing { get; set; } = 1.15;

    /// <summary>
    /// Colour of captions ("Figure 1 – …").
    /// </summary>
    public string CaptionColor { get; set; } = "#1F4E79";

    /// <summary>
    /// Fill of a table's header row.
    /// </summary>
    public string TableHeaderFill { get; set; } = "#BDD7EE";

    /// <summary>
    /// Text colour of a table's header row.
    /// </summary>
    public string TableHeaderColor { get; set; } = "#1F3864";

    /// <summary>
    /// Fill of every other body row of a table.
    /// </summary>
    public string TableBandFill { get; set; } = "#EEF5FB";

    /// <summary>
    /// Colour of table lines.
    /// </summary>
    public string TableBorderColor { get; set; } = "#9DC3E6";

    /// <summary>
    /// Label of figure captions, for example "Şekil" in a Turkish document.
    /// </summary>
    public string FigureLabel { get; set; } = "Figure";

    /// <summary>
    /// Label of table captions, for example "Tablo".
    /// </summary>
    public string TableLabel { get; set; } = "Table";

    /// <summary>
    /// What stands between a caption's number and its text.
    /// </summary>
    public string CaptionSeparator { get; set; } = " – ";

    internal string Heading => HeadingFont ?? BodyFont;

    internal double HeadingSize(int level) =>
        HeadingSizes.Count == 0 ? BodySize : HeadingSizes[System.Math.Min(level, HeadingSizes.Count) - 1];
}
