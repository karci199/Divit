#nullable enable
namespace Divit.Docx;

using Divit.Packaging;
using System;
using System.Xml;

/// <content>
/// Styles, the heading numbering and the document settings.
/// </content>
public sealed partial class DocxDocument
{
    /// <summary>
    /// The styles, built from the look: document defaults, Normal, headings, title, subtitle, caption, no spacing,
    /// and the table of contents styles. Built-in styles carry Word's own names, which Word shows translated.
    /// </summary>
    private void WriteStyles(XmlWriter w, Parts parts)
    {
        var look = Look;
        w.WriteStartElement("w", "styles", Main);

        w.WriteStartElement("docDefaults", Main);
        w.WriteStartElement("rPrDefault", Main);
        w.WriteStartElement("rPr", Main);
        Fonts(w, look.BodyFont);
        Size(w, look.BodySize);
        if (Language != null)
        {
            w.WriteStartElement("lang", Main);
            w.WriteAttributeString("w", "val", Main, Language);
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("pPrDefault", Main);
        w.WriteStartElement("pPr", Main);
        Spacing(w, null, look.SpaceAfter, look.LineSpacing);
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();

        StyleStart(w, "paragraph", DocxStyle.Normal, "Normal", null, null, 0, true, true);
        w.WriteStartElement("rPr", Main);
        Val(w, "color", OfficeText.Rgb(look.BodyColor));
        w.WriteEndElement();
        w.WriteEndElement();

        StyleStart(w, "character", "DefaultParagraphFont", "Default Paragraph Font", null, null, 1, false, true, hidden: true);
        w.WriteEndElement();

        // The default table style: Word needs it once the document has a table.
        StyleStart(w, "table", "TableNormal", "Normal Table", null, null, 99, false, true, hidden: true);
        w.WriteStartElement("tblPr", Main);
        w.WriteStartElement("tblInd", Main);
        w.WriteAttributeString("w", "w", Main, "0");
        w.WriteAttributeString("w", "type", Main, "dxa");
        w.WriteEndElement();
        w.WriteStartElement("tblCellMar", Main);
        foreach (var (side, width) in new[] { ("top", "0"), ("left", "108"), ("bottom", "0"), ("right", "108") })
        {
            w.WriteStartElement(side, Main);
            w.WriteAttributeString("w", "w", Main, width);
            w.WriteAttributeString("w", "type", Main, "dxa");
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();

        StyleStart(w, "numbering", "NoList", "No List", null, null, 99, false, true, hidden: true);
        w.WriteEndElement();

        for (var level = 1; level <= 6; level++)
        {
            StyleStart(w, "paragraph", DocxStyle.Heading(level), "heading " + Int(level), DocxStyle.Normal, DocxStyle.Normal, 9, true, false);
            w.WriteStartElement("pPr", Main);
            w.WriteElementString("keepNext", Main, null);
            w.WriteElementString("keepLines", Main, null);
            if (NumberHeadings)
            {
                NumberingProperties(w, level - 1, 1);
            }

            Spacing(w, level == 1 ? 18 : 12, 6, null);
            Val(w, "outlineLvl", Int(level - 1));
            w.WriteEndElement();
            HeadingRun(w, look.HeadingSize(level), look.HeadingColor, level <= 3);
            w.WriteEndElement();
        }

        StyleStart(w, "paragraph", DocxStyle.Title, "Title", DocxStyle.Normal, DocxStyle.Normal, 10, true, false);
        w.WriteStartElement("pPr", Main);
        Spacing(w, null, 12, 1);
        w.WriteEndElement();
        HeadingRun(w, look.TitleSize, look.HeadingColor, true);
        w.WriteEndElement();

        StyleStart(w, "paragraph", DocxStyle.Subtitle, "Subtitle", DocxStyle.Normal, DocxStyle.Normal, 11, true, false);
        w.WriteStartElement("pPr", Main);
        Spacing(w, null, 12, null);
        w.WriteEndElement();
        HeadingRun(w, look.BodySize + 3, look.HeadingColor, false);
        w.WriteEndElement();

        StyleStart(w, "paragraph", DocxStyle.Caption, "caption", DocxStyle.Normal, DocxStyle.Normal, 35, true, false, unhide: true);
        w.WriteStartElement("pPr", Main);
        Spacing(w, 3, 9, 1);
        w.WriteEndElement();
        w.WriteStartElement("rPr", Main);
        w.WriteElementString("i", Main, null);
        w.WriteElementString("iCs", Main, null);
        Val(w, "color", OfficeText.Rgb(look.CaptionColor));
        Size(w, Math.Max(8, look.BodySize - 1));
        w.WriteEndElement();
        w.WriteEndElement();

        StyleStart(w, "paragraph", DocxStyle.NoSpacing, "No Spacing", null, null, 1, true, false);
        w.WriteStartElement("pPr", Main);
        Spacing(w, 0, 0, 1);
        w.WriteEndElement();
        w.WriteEndElement();

        // The title of the table of contents looks like a first-level heading but is neither numbered nor listed.
        StyleStart(w, "paragraph", TocHeadingStyle, "TOC Heading", DocxStyle.Heading(1), DocxStyle.Normal, 39, true, false, unhide: true);
        w.WriteStartElement("pPr", Main);
        if (NumberHeadings)
        {
            NumberingProperties(w, 0, 0);
        }

        Val(w, "outlineLvl", "9");
        w.WriteEndElement();
        w.WriteEndElement();

        for (var level = 1; level <= 6; level++)
        {
            StyleStart(w, "paragraph", "TOC" + Int(level), "toc " + Int(level), DocxStyle.Normal, DocxStyle.Normal, 39, false, false, unhide: true);
            w.WriteStartElement("pPr", Main);
            Spacing(w, null, level == 1 ? 6 : 3, null);
            if (level > 1)
            {
                w.WriteStartElement("ind", Main);
                w.WriteAttributeString("w", "left", Main, Int(220 * (level - 1)));
                w.WriteEndElement();
            }

            w.WriteEndElement();
            if (level == 1)
            {
                w.WriteStartElement("rPr", Main);
                w.WriteElementString("b", Main, null);
                w.WriteElementString("bCs", Main, null);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        StyleStart(w, "paragraph", FiguresStyle, "table of figures", DocxStyle.Normal, DocxStyle.Normal, 99, false, false, unhide: true);
        w.WriteStartElement("pPr", Main);
        Spacing(w, null, 3, null);
        w.WriteEndElement();
        w.WriteEndElement();
        WriteFeatureStyles(w, parts);

        TableStyle(w, "DivitTable", "Divit Table", true, true);
        TableStyle(w, "DivitGrid", "Divit Grid", true, false);
        TableStyle(w, "DivitPlain", "Divit Plain", false, false);
        w.WriteEndElement();
    }

    /// <summary>
    /// The styles only some documents need, written when used: list items, links, footnotes.
    /// </summary>
    private void WriteFeatureStyles(XmlWriter w, Parts parts)
    {
        if (parts.HasLists)
        {
            StyleStart(w, "paragraph", DocxStyle.ListParagraph, "List Paragraph", DocxStyle.Normal, null, 34, true, false);
            w.WriteStartElement("pPr", Main);
            w.WriteStartElement("ind", Main);
            w.WriteAttributeString("w", "left", Main, "720");
            w.WriteEndElement();
            w.WriteElementString("contextualSpacing", Main, null);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        if (parts.HasLinks)
        {
            StyleStart(w, "character", LinkStyle, "Hyperlink", null, null, 99, false, false, unhide: true);
            w.WriteStartElement("rPr", Main);
            Val(w, "color", OfficeText.Rgb(Look.LinkColor));
            Val(w, "u", "single");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        if (parts.Footnotes.Count > 0)
        {
            StyleStart(w, "paragraph", FootnoteTextStyle, "footnote text", DocxStyle.Normal, null, 99, false, false, unhide: true);
            w.WriteStartElement("pPr", Main);
            Spacing(w, 0, 0, 1);
            w.WriteEndElement();
            w.WriteStartElement("rPr", Main);
            Size(w, Math.Max(7, Look.BodySize - 2));
            w.WriteEndElement();
            w.WriteEndElement();

            StyleStart(w, "character", FootnoteReferenceStyle, "footnote reference", null, null, 99, false, false, unhide: true);
            w.WriteStartElement("rPr", Main);
            Val(w, "vertAlign", "superscript");
            w.WriteEndElement();
            w.WriteEndElement();
        }
    }

    /// <summary>
    /// A table style from the look: compact paragraphs, cell margins, lines (or none), and for the house table a
    /// coloured bold header row and banded body rows. Word lists it under Table Design, so readers can switch.
    /// </summary>
    private void TableStyle(XmlWriter w, string id, string name, bool lines, bool banded)
    {
        var look = Look;
        StyleStart(w, "table", id, name, "TableNormal", null, 59, false, false);
        w.WriteStartElement("pPr", Main);
        Spacing(w, 2, 2, 1);
        w.WriteEndElement();
        w.WriteStartElement("tblPr", Main);
        if (banded)
        {
            Val(w, "tblStyleRowBandSize", "1");
        }

        if (lines)
        {
            var border = new DocxBorder(look.TableBorderColor, 0.5);
            w.WriteStartElement("tblBorders", Main);
            foreach (var side in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
            {
                Border(w, side, border, 0);
            }

            w.WriteEndElement();
        }

        w.WriteStartElement("tblCellMar", Main);
        foreach (var (side, width) in new[] { ("top", "57"), ("left", "108"), ("bottom", "57"), ("right", "108") })
        {
            w.WriteStartElement(side, Main);
            w.WriteAttributeString("w", "w", Main, width);
            w.WriteAttributeString("w", "type", Main, "dxa");
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteEndElement();

        if (lines)
        {
            // The first row: bold in the header colours; a band of colour on every other body row.
            w.WriteStartElement("tblStylePr", Main);
            w.WriteAttributeString("w", "type", Main, "firstRow");
            w.WriteStartElement("rPr", Main);
            w.WriteElementString("b", Main, null);
            w.WriteElementString("bCs", Main, null);
            Val(w, "color", OfficeText.Rgb(banded ? look.TableHeaderColor : look.BodyColor));
            w.WriteEndElement();
            w.WriteStartElement("tcPr", Main);
            Shading(w, banded ? look.TableHeaderFill : look.TableBandFill);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        if (banded)
        {
            w.WriteStartElement("tblStylePr", Main);
            w.WriteAttributeString("w", "type", Main, "band1Horz");
            w.WriteStartElement("tcPr", Main);
            Shading(w, look.TableBandFill);
            w.WriteEndElement();
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Opens a style with its name and the flags Word reads, in the order of the standard; the caller adds the
    /// properties and closes it.
    /// </summary>
    private static void StyleStart(XmlWriter w, string type, string id, string name, string? basedOn, string? next, int priority, bool quick, bool isDefault, bool hidden = false, bool unhide = false)
    {
        w.WriteStartElement("style", Main);
        w.WriteAttributeString("w", "type", Main, type);
        if (isDefault) { w.WriteAttributeString("w", "default", Main, "1"); }
        w.WriteAttributeString("w", "styleId", Main, id);
        Val(w, "name", name);
        if (basedOn != null) { Val(w, "basedOn", basedOn); }
        if (next != null) { Val(w, "next", next); }
        if (priority > 0) { Val(w, "uiPriority", Int(priority)); }
        if (hidden) { w.WriteElementString("semiHidden", Main, null); }
        if (hidden || unhide) { w.WriteElementString("unhideWhenUsed", Main, null); }
        if (quick) { w.WriteElementString("qFormat", Main, null); }
    }

    private void HeadingRun(XmlWriter w, double size, string color, bool bold)
    {
        w.WriteStartElement("rPr", Main);
        if (Look.HeadingFont != null) { Fonts(w, Look.HeadingFont); }
        if (bold) { w.WriteElementString("b", Main, null); w.WriteElementString("bCs", Main, null); }
        Val(w, "color", OfficeText.Rgb(color));
        Size(w, size);
        w.WriteEndElement();
    }

    /// <summary>
    /// Space above and below in points and line spacing as a multiple; empty values are left out.
    /// </summary>
    private static void Spacing(XmlWriter w, double? before, double? after, double? line)
    {
        w.WriteStartElement("spacing", Main);
        if (before is double b) { w.WriteAttributeString("w", "before", Main, Int(DocxUnits.PointTwips(b))); }
        if (after is double a) { w.WriteAttributeString("w", "after", Main, Int(DocxUnits.PointTwips(a))); }
        if (line is double l)
        {
            w.WriteAttributeString("w", "line", Main, Int((int)Math.Round(l * 240)));
            w.WriteAttributeString("w", "lineRule", Main, "auto");
        }

        w.WriteEndElement();
    }

    private static void NumberingProperties(XmlWriter w, int level, int number)
    {
        w.WriteStartElement("numPr", Main);
        Val(w, "ilvl", Int(level));
        Val(w, "numId", Int(number));
        w.WriteEndElement();
    }


    /// <summary>
    /// The numbering part: the heading outline (1, 1.1, 1.1.1 … on the heading styles), then a definition per list
    /// kind, then one numbering per list. Every numbered list starts again at 1.
    /// </summary>
    private void WriteNumbering(XmlWriter w, Parts parts)
    {
        w.WriteStartElement("w", "numbering", Main);
        if (NumberHeadings)
        {
            w.WriteStartElement("abstractNum", Main);
            w.WriteAttributeString("w", "abstractNumId", Main, "0");
            Val(w, "multiLevelType", "multilevel");
            for (var level = 0; level < 6; level++)
            {
                w.WriteStartElement("lvl", Main);
                w.WriteAttributeString("w", "ilvl", Main, Int(level));
                Val(w, "start", "1");
                Val(w, "numFmt", "decimal");
                Val(w, "pStyle", DocxStyle.Heading(level + 1));
                var text = new string[level + 1];
                for (var l = 0; l <= level; l++)
                {
                    text[l] = "%" + Int(l + 1);
                }

                Val(w, "lvlText", string.Join(".", text));
                Val(w, "lvlJc", "left");
                Indent(w, 432 + (144 * level), 432 + (144 * level));
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        foreach (var kind in parts.ListKinds)
        {
            w.WriteStartElement("abstractNum", Main);
            w.WriteAttributeString("w", "abstractNumId", Main, Int(ListDefinition(kind)));
            Val(w, "multiLevelType", "hybridMultilevel");
            for (var level = 0; level < 9; level++)
            {
                w.WriteStartElement("lvl", Main);
                w.WriteAttributeString("w", "ilvl", Main, Int(level));
                Val(w, "start", "1");
                if (kind == DocxListKind.Bullet)
                {
                    Val(w, "numFmt", "bullet");
                    Val(w, "lvlText", level % 3 == 0 ? "\u2022" : level % 3 == 1 ? "\u2013" : "\u00BB");
                }
                else
                {
                    Val(w, "numFmt", level % 3 == 0 ? "decimal" : level % 3 == 1 ? "lowerLetter" : "lowerRoman");
                    Val(w, "lvlText", "%" + Int(level + 1) + (level % 3 == 1 ? ")" : "."));
                }

                Val(w, "lvlJc", "left");
                Indent(w, 720 * (level + 1), 360);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        if (NumberHeadings)
        {
            w.WriteStartElement("num", Main);
            w.WriteAttributeString("w", "numId", Main, "1");
            Val(w, "abstractNumId", "0");
            w.WriteEndElement();
        }

        foreach (var list in parts.ListIds)
        {
            w.WriteStartElement("num", Main);
            w.WriteAttributeString("w", "numId", Main, Int(list.Value));
            Val(w, "abstractNumId", Int(ListDefinition(list.Key.Kind)));
            if (list.Key.Kind == DocxListKind.Numbered)
            {
                // Lists of one definition would count on from each other without this.
                w.WriteStartElement("lvlOverride", Main);
                w.WriteAttributeString("w", "ilvl", Main, "0");
                Val(w, "startOverride", "1");
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// The definition id of a list kind; 0 is the heading outline.
    /// </summary>
    private static int ListDefinition(DocxListKind kind) => kind == DocxListKind.Bullet ? 1 : 2;

    /// <summary>
    /// The indent of a numbering level: the text at <paramref name="left"/>, the number hanging before it.
    /// </summary>
    private static void Indent(XmlWriter w, int left, int hanging)
    {
        w.WriteStartElement("pPr", Main);
        w.WriteStartElement("ind", Main);
        w.WriteAttributeString("w", "left", Main, Int(left));
        w.WriteAttributeString("w", "hanging", Main, Int(hanging));
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// Settings in the order of the standard: tracking, protection, tabs, even pages, field updates, footnote
    /// separators; Word 2013+ compatibility, so Word does not open the file in compatibility mode.
    /// </summary>
    private void WriteSettings(XmlWriter w, Parts parts)
    {
        w.WriteStartElement("w", "settings", Main);
        if (TrackChanges) { w.WriteElementString("trackRevisions", Main, null); }
        if (Protection is DocxProtection protection)
        {
            w.WriteStartElement("documentProtection", Main);
            w.WriteAttributeString("w", "edit", Main, protection == DocxProtection.Forms ? "forms" : protection == DocxProtection.Comments ? "comments" : protection == DocxProtection.TrackedChanges ? "trackedChanges" : "readOnly");
            w.WriteAttributeString("w", "enforcement", Main, "1");
            if (Password != null)
            {
                var salt = PasswordHash.Salt(Created == null ? null : "Divit|document|" + Created.Value.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture));
                w.WriteAttributeString("w", "algorithmName", Main, PasswordHash.Algorithm);
                w.WriteAttributeString("w", "hashValue", Main, PasswordHash.Sha512(Password, salt));
                w.WriteAttributeString("w", "saltValue", Main, Convert.ToBase64String(salt));
                w.WriteAttributeString("w", "spinCount", Main, Int(PasswordHash.SpinCount));
            }

            w.WriteEndElement();
        }

        Val(w, "defaultTabStop", "708");
        if (parts.EvenPages) { w.WriteElementString("evenAndOddHeaders", Main, null); }
        Val(w, "characterSpacingControl", "doNotCompress");
        if (parts.HasContents || parts.HasReferences) { Val(w, "updateFields", "true"); }
        if (parts.Footnotes.Count > 0)
        {
            w.WriteStartElement("footnotePr", Main);
            foreach (var id in new[] { "-1", "0" })
            {
                w.WriteStartElement("footnote", Main);
                w.WriteAttributeString("w", "id", Main, id);
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        w.WriteStartElement("compat", Main);
        w.WriteStartElement("compatSetting", Main);
        w.WriteAttributeString("w", "name", Main, "compatibilityMode");
        w.WriteAttributeString("w", "uri", Main, "http://schemas.microsoft.com/office/word");
        w.WriteAttributeString("w", "val", Main, "15");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }
}
