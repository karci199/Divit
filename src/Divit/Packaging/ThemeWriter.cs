#nullable enable
namespace Divit.Packaging;

using System.Xml;

/// <summary>
/// Writes a DrawingML theme part (<c>theme1.xml</c>), the same in every Office format: colour scheme, font scheme and
/// the smallest format scheme the standard allows (three plain fills, lines, effects and backgrounds).
/// </summary>
internal static class ThemeWriter
{
    /// <summary>
    /// Content type of a theme part.
    /// </summary>
    public const string ContentType = "application/vnd.openxmlformats-officedocument.theme+xml";

    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static void Write(XmlWriter w, OfficeTheme theme, string defaultFont)
    {
        w.WriteStartElement("a", "theme", A);
        w.WriteAttributeString("name", OfficeText.Clean(theme.Name));
        w.WriteStartElement("a", "themeElements", A);

        w.WriteStartElement("a", "clrScheme", A);
        w.WriteAttributeString("name", OfficeText.Clean(theme.Name));
        Color(w, "dk1", theme.Dark1);
        Color(w, "lt1", theme.Light1);
        Color(w, "dk2", theme.Dark2);
        Color(w, "lt2", theme.Light2);
        Color(w, "accent1", theme.Accent1);
        Color(w, "accent2", theme.Accent2);
        Color(w, "accent3", theme.Accent3);
        Color(w, "accent4", theme.Accent4);
        Color(w, "accent5", theme.Accent5);
        Color(w, "accent6", theme.Accent6);
        Color(w, "hlink", theme.Hyperlink);
        Color(w, "folHlink", theme.FollowedHyperlink);
        w.WriteEndElement();

        w.WriteStartElement("a", "fontScheme", A);
        w.WriteAttributeString("name", OfficeText.Clean(theme.Name));
        Font(w, "majorFont", theme.HeadingFont ?? defaultFont);
        Font(w, "minorFont", theme.BodyFont ?? defaultFont);
        w.WriteEndElement();

        w.WriteStartElement("a", "fmtScheme", A);
        w.WriteAttributeString("name", OfficeText.Clean(theme.Name));
        w.WriteStartElement("a", "fillStyleLst", A);
        for (var i = 0; i < 3; i++) { PlaceholderFill(w); }
        w.WriteEndElement();
        w.WriteStartElement("a", "lnStyleLst", A);
        foreach (var width in new[] { "6350", "12700", "19050" })
        {
            w.WriteStartElement("a", "ln", A);
            w.WriteAttributeString("w", width);
            PlaceholderFill(w);
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteStartElement("a", "effectStyleLst", A);
        for (var i = 0; i < 3; i++)
        {
            w.WriteStartElement("a", "effectStyle", A);
            w.WriteElementString("a", "effectLst", A, string.Empty);
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteStartElement("a", "bgFillStyleLst", A);
        for (var i = 0; i < 3; i++) { PlaceholderFill(w); }
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteEndElement();
        w.WriteElementString("a", "objectDefaults", A, string.Empty);
        w.WriteElementString("a", "extraClrSchemeLst", A, string.Empty);
        w.WriteEndElement();
    }

    private static void Color(XmlWriter w, string name, string color)
    {
        w.WriteStartElement("a", name, A);
        w.WriteStartElement("a", "srgbClr", A);
        w.WriteAttributeString("val", OfficeText.Rgb(color));
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void Font(XmlWriter w, string name, string typeface)
    {
        w.WriteStartElement("a", name, A);
        foreach (var script in new[] { "latin", "ea", "cs" })
        {
            w.WriteStartElement("a", script, A);
            w.WriteAttributeString("typeface", script == "latin" ? OfficeText.Clean(typeface) : string.Empty);
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// A solid fill in the placeholder colour: the style takes the colour of whatever uses it.
    /// </summary>
    private static void PlaceholderFill(XmlWriter w)
    {
        w.WriteStartElement("a", "solidFill", A);
        w.WriteStartElement("a", "schemeClr", A);
        w.WriteAttributeString("val", "phClr");
        w.WriteEndElement();
        w.WriteEndElement();
    }
}
