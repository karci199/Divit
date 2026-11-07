#nullable enable
namespace Divit.Xlsx;

/// <summary>
/// Excel's built-in cell styles, the ones in the Cell Styles gallery. A cell with one of them shows it highlighted in
/// the gallery, and its look follows Excel's default Office theme.
/// </summary>
public enum XlsxCellStyle
{
    /// <summary>
    /// The default style.
    /// </summary>
    Normal,

    /// <summary>
    /// Large title text.
    /// </summary>
    Title,

    /// <summary>
    /// First heading level: large, bold, thick accent line below.
    /// </summary>
    Heading1,

    /// <summary>
    /// Second heading level.
    /// </summary>
    Heading2,

    /// <summary>
    /// Third heading level.
    /// </summary>
    Heading3,

    /// <summary>
    /// Fourth heading level: bold only.
    /// </summary>
    Heading4,

    /// <summary>
    /// Green: a good value.
    /// </summary>
    Good,

    /// <summary>
    /// Red: a bad value.
    /// </summary>
    Bad,

    /// <summary>
    /// Yellow: a neutral value.
    /// </summary>
    Neutral,

    /// <summary>
    /// Orange box: a value the user types in.
    /// </summary>
    Input,

    /// <summary>
    /// Grey box, bold: a result.
    /// </summary>
    Output,

    /// <summary>
    /// Grey box, orange bold: a computed value.
    /// </summary>
    Calculation,

    /// <summary>
    /// Dark grey box, white bold, double frame: a check.
    /// </summary>
    CheckCell,

    /// <summary>
    /// Orange text, double line below: a value taken from another cell.
    /// </summary>
    LinkedCell,

    /// <summary>
    /// Pale yellow box: a remark.
    /// </summary>
    Note,

    /// <summary>
    /// Red text: a warning.
    /// </summary>
    WarningText,

    /// <summary>
    /// Grey italic text: an explanation.
    /// </summary>
    Explanatory,

    /// <summary>
    /// Bold, thin line above and double line below: a total.
    /// </summary>
    Total,
}

/// <summary>
/// Names, built-in ids and looks of <see cref="XlsxCellStyle"/>.
/// </summary>
internal static class XlsxCellStyles
{
    public static string Name(XlsxCellStyle style)
    {
        switch (style)
        {
            case XlsxCellStyle.Heading1: return "Heading 1";
            case XlsxCellStyle.Heading2: return "Heading 2";
            case XlsxCellStyle.Heading3: return "Heading 3";
            case XlsxCellStyle.Heading4: return "Heading 4";
            case XlsxCellStyle.CheckCell: return "Check Cell";
            case XlsxCellStyle.LinkedCell: return "Linked Cell";
            case XlsxCellStyle.WarningText: return "Warning Text";
            case XlsxCellStyle.Explanatory: return "Explanatory Text";
            default: return style.ToString();
        }
    }

    /// <summary>
    /// The builtinId Excel knows the style by (ECMA-376 part 1, 18.8.7).
    /// </summary>
    public static int BuiltinId(XlsxCellStyle style)
    {
        switch (style)
        {
            case XlsxCellStyle.Title: return 15;
            case XlsxCellStyle.Heading1: return 16;
            case XlsxCellStyle.Heading2: return 17;
            case XlsxCellStyle.Heading3: return 18;
            case XlsxCellStyle.Heading4: return 19;
            case XlsxCellStyle.Input: return 20;
            case XlsxCellStyle.Output: return 21;
            case XlsxCellStyle.Calculation: return 22;
            case XlsxCellStyle.CheckCell: return 23;
            case XlsxCellStyle.LinkedCell: return 24;
            case XlsxCellStyle.Total: return 25;
            case XlsxCellStyle.Good: return 26;
            case XlsxCellStyle.Bad: return 27;
            case XlsxCellStyle.Neutral: return 28;
            case XlsxCellStyle.Note: return 10;
            case XlsxCellStyle.WarningText: return 11;
            case XlsxCellStyle.Explanatory: return 53;
            default: return 0;
        }
    }

    /// <summary>
    /// The font, fill and borders of a style, with Excel's default Office theme colours.
    /// </summary>
    public static XlsxStyle Look(XlsxCellStyle style)
    {
        var thin = new XlsxBorder("thin", "#7F7F7F");
        switch (style)
        {
            case XlsxCellStyle.Title:
                return new XlsxStyle { FontSize = 18, FontColor = "#44546A" };
            case XlsxCellStyle.Heading1:
                return new XlsxStyle { FontSize = 15, Bold = true, FontColor = "#44546A", Bottom = new XlsxBorder("thick", "#4472C4") };
            case XlsxCellStyle.Heading2:
                return new XlsxStyle { FontSize = 13, Bold = true, FontColor = "#44546A", Bottom = new XlsxBorder("thick", "#A2B8E1") };
            case XlsxCellStyle.Heading3:
                return new XlsxStyle { Bold = true, FontColor = "#44546A", Bottom = new XlsxBorder("medium", "#8EA9DB") };
            case XlsxCellStyle.Heading4:
                return new XlsxStyle { Bold = true, FontColor = "#44546A" };
            case XlsxCellStyle.Good:
                return new XlsxStyle { Fill = "#C6EFCE", FontColor = "#006100" };
            case XlsxCellStyle.Bad:
                return new XlsxStyle { Fill = "#FFC7CE", FontColor = "#9C0006" };
            case XlsxCellStyle.Neutral:
                return new XlsxStyle { Fill = "#FFEB9C", FontColor = "#9C5700" };
            case XlsxCellStyle.Input:
                return Boxed(new XlsxStyle { Fill = "#FFCC99", FontColor = "#3F3F76" }, thin);
            case XlsxCellStyle.Output:
                return Boxed(new XlsxStyle { Fill = "#F2F2F2", FontColor = "#3F3F3F", Bold = true }, new XlsxBorder("thin", "#3F3F3F"));
            case XlsxCellStyle.Calculation:
                return Boxed(new XlsxStyle { Fill = "#F2F2F2", FontColor = "#FA7D00", Bold = true }, thin);
            case XlsxCellStyle.CheckCell:
                return Boxed(new XlsxStyle { Fill = "#A5A5A5", FontColor = "#FFFFFF", Bold = true }, new XlsxBorder("double", "#3F3F3F"));
            case XlsxCellStyle.LinkedCell:
                return new XlsxStyle { FontColor = "#FA7D00", Bottom = new XlsxBorder("double", "#FF8001") };
            case XlsxCellStyle.Note:
                return Boxed(new XlsxStyle { Fill = "#FFFFCC" }, new XlsxBorder("thin", "#B2B2B2"));
            case XlsxCellStyle.WarningText:
                return new XlsxStyle { FontColor = "#FF0000" };
            case XlsxCellStyle.Explanatory:
                return new XlsxStyle { Italic = true, FontColor = "#7F7F7F" };
            case XlsxCellStyle.Total:
                return new XlsxStyle { Bold = true, Top = new XlsxBorder("thin", "#4472C4"), Bottom = new XlsxBorder("double", "#4472C4") };
            default:
                return default;
        }
    }

    private static XlsxStyle Boxed(XlsxStyle style, XlsxBorder border)
    {
        style.Left = border;
        style.Right = border;
        style.Top = border;
        style.Bottom = border;
        return style;
    }
}
