#nullable enable
namespace Divit.Docx;

using Divit.Packaging;
using System;
using System.Globalization;
using System.Xml;

/// <content>
/// Cross-references, check boxes, footnotes and watermarks.
/// </content>
public sealed partial class DocxDocument
{
    private const string LinkStyle = "Hyperlink";
    private const string FootnoteTextStyle = "FootnoteText";
    private const string FootnoteReferenceStyle = "FootnoteReference";
    private const string CompatibilityNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string Word2010Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string VmlNs = "urn:schemas-microsoft-com:vml";
    private const string OfficeVmlNs = "urn:schemas-microsoft-com:office:office";
    private const string WordVmlNs = "urn:schemas-microsoft-com:office:word";

    /// <summary>
    /// The box characters of a check box, ticked and empty, in the font Word uses for them.
    /// </summary>
    private const string CheckBoxFont = "MS Gothic";

    /// <summary>
    /// A cross-reference: a REF field on the target's bookmark (its text, or with <c>\r</c> its paragraph number) or
    /// a PAGEREF field, clickable with <c>\h</c>. It shows what Divit knows until Word updates it.
    /// </summary>
    private static void WriteReference(XmlWriter w, DocxRun run, DocxParagraph target, Parts parts)
    {
        string code, result;
        switch (run.ReferenceKind)
        {
            case DocxReferenceKind.Label:
                code = " REF " + BookmarkName(parts.LabelBookmarks[target]) + " \\h ";
                result = target.CaptionLabel ?? string.Empty;
                break;
            case DocxReferenceKind.Number:
                code = " REF " + BookmarkName(parts.Bookmarks[target]) + " \\r \\h ";
                result = parts.HeadingNumbers.TryGetValue(target, out var number) ? number : string.Empty;
                break;
            case DocxReferenceKind.Page:
                code = " PAGEREF " + BookmarkName(parts.Bookmarks[target]) + " \\h ";
                result = "1";
                break;
            default:
                code = " REF " + BookmarkName(parts.Bookmarks[target]) + " \\h ";
                result = target.Text;
                break;
        }

        w.WriteStartElement("fldSimple", Main);
        w.WriteAttributeString("w", "instr", Main, code);
        w.WriteStartElement("r", Main);
        WriteRunProperties(w, run.Font);
        Text(w, result);
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// A check box content control (Word 2010): the control's state, and the box character it shows.
    /// </summary>
    private static void WriteCheckBox(XmlWriter w, DocxRun run, bool ticked, Parts parts)
    {
        w.WriteStartElement("sdt", Main);
        w.WriteStartElement("sdtPr", Main);
        Val(w, "id", Int(++parts.Controls));
        w.WriteStartElement("w14", "checkbox", Word2010Ns);
        w.WriteStartElement("w14", "checked", Word2010Ns);
        w.WriteAttributeString("w14", "val", Word2010Ns, ticked ? "1" : "0");
        w.WriteEndElement();
        foreach (var (name, glyph) in new[] { ("checkedState", "2612"), ("uncheckedState", "2610") })
        {
            w.WriteStartElement("w14", name, Word2010Ns);
            w.WriteAttributeString("w14", "val", Word2010Ns, glyph);
            w.WriteAttributeString("w14", "font", Word2010Ns, CheckBoxFont);
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("sdtContent", Main);
        w.WriteStartElement("r", Main);
        var font = run.Font ?? default;
        font.Name = CheckBoxFont;
        WriteRunProperties(w, font);
        Text(w, ticked ? "☒" : "☐");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The footnotes part: the separator lines Word draws above the notes, then the notes, each starting with its
    /// number.
    /// </summary>
    private static void WriteFootnotes(XmlWriter w, Parts parts)
    {
        w.WriteStartElement("w", "footnotes", Main);
        foreach (var (type, id) in new[] { ("separator", "-1"), ("continuationSeparator", "0") })
        {
            w.WriteStartElement("footnote", Main);
            w.WriteAttributeString("w", "type", Main, type);
            w.WriteAttributeString("w", "id", Main, id);
            w.WriteStartElement("p", Main);
            w.WriteStartElement("pPr", Main);
            Spacing(w, null, 0, 1);
            w.WriteEndElement();
            w.WriteStartElement("r", Main);
            w.WriteElementString(type, Main, null);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        foreach (var note in parts.Footnotes)
        {
            w.WriteStartElement("footnote", Main);
            w.WriteAttributeString("w", "id", Main, Int(parts.FootnoteIds[note]));
            w.WriteStartElement("p", Main);
            w.WriteStartElement("pPr", Main);
            Val(w, "pStyle", FootnoteTextStyle);
            w.WriteEndElement();
            w.WriteStartElement("r", Main);
            WriteRunProperties(w, null, FootnoteReferenceStyle);
            w.WriteElementString("footnoteRef", Main, null);
            w.WriteEndElement();
            w.WriteStartElement("r", Main);
            Text(w, " " + note.Footnote);
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// A watermark as Word draws one: slanted WordArt text (a VML text path) centred on the page behind the text, in
    /// a paragraph of its own at the top of the header.
    /// </summary>
    private void WriteWatermark(XmlWriter w, DocxSection section, Parts parts)
    {
        var text = section.Watermark!;
        var width = section.TextWidthTwips / 20.0 * 0.9;
        var height = width * 3 / (text.Length + 2.5);
        var id = ++parts.Shapes;

        w.WriteStartElement("p", Main);
        w.WriteStartElement("r", Main);
        w.WriteStartElement("pict", Main);
        // Each header is a part of its own with one watermark, so each declares the shape type.
        WriteTextShapeType(w);

        w.WriteStartElement("v", "shape", VmlNs);
        w.WriteAttributeString("id", "DivitWatermark" + Int(id));
        w.WriteAttributeString("o", "spid", OfficeVmlNs, "_x0000_s" + Int(2048 + id));
        w.WriteAttributeString("type", "#_x0000_t136");
        w.WriteAttributeString(
            "style",
            "position:absolute;margin-left:0;margin-top:0;width:" + Points(width) + ";height:" + Points(height)
            + ";rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;"
            + "mso-position-vertical:center;mso-position-vertical-relative:margin");
        w.WriteAttributeString("o", "allowincell", OfficeVmlNs, "f");
        w.WriteAttributeString("fillcolor", "#" + OfficeText.Rgb(section.WatermarkColor));
        w.WriteAttributeString("stroked", "f");
        w.WriteStartElement("v", "fill", VmlNs);
        w.WriteAttributeString("opacity", ".5");
        w.WriteEndElement();
        w.WriteStartElement("v", "textpath", VmlNs);
        w.WriteAttributeString("style", "font-family:\"" + Look.BodyFont + "\";font-size:1pt");
        w.WriteAttributeString("string", OfficeText.Clean(text));
        w.WriteEndElement();
        w.WriteStartElement("w10", "wrap", WordVmlNs);
        w.WriteAttributeString("anchorx", "margin");
        w.WriteAttributeString("anchory", "margin");
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The shape type of WordArt text along a straight path (type 136), as Word writes it before the first use.
    /// </summary>
    private static void WriteTextShapeType(XmlWriter w)
    {
        w.WriteStartElement("v", "shapetype", VmlNs);
        w.WriteAttributeString("id", "_x0000_t136");
        w.WriteAttributeString("coordsize", "21600,21600");
        w.WriteAttributeString("o", "spt", OfficeVmlNs, "136");
        w.WriteAttributeString("adj", "10800");
        w.WriteAttributeString("path", "m@7,l@8,m@5,21600l@6,21600e");
        w.WriteStartElement("v", "formulas", VmlNs);
        foreach (var equation in new[]
        {
            "sum #0 0 10800", "prod #0 2 1", "sum 21600 0 @1", "sum 0 0 @2", "sum 21600 0 @3", "if @0 @3 0",
            "if @0 21600 @1", "if @0 0 @2", "if @0 @4 21600", "mid @5 @6", "mid @8 @5", "mid @7 @8", "mid @6 @7",
            "sum @6 0 @5",
        })
        {
            w.WriteStartElement("v", "f", VmlNs);
            w.WriteAttributeString("eqn", equation);
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteStartElement("v", "path", VmlNs);
        w.WriteAttributeString("textpathok", "t");
        w.WriteAttributeString("o", "connecttype", OfficeVmlNs, "custom");
        w.WriteAttributeString("o", "connectlocs", OfficeVmlNs, "@9,0;@10,10800;@11,21600;@12,10800");
        w.WriteAttributeString("o", "connectangles", OfficeVmlNs, "270,180,90,0");
        w.WriteEndElement();
        w.WriteStartElement("v", "textpath", VmlNs);
        w.WriteAttributeString("on", "t");
        w.WriteAttributeString("fitshape", "t");
        w.WriteEndElement();
        w.WriteStartElement("v", "handles", VmlNs);
        w.WriteStartElement("v", "h", VmlNs);
        w.WriteAttributeString("position", "#0,bottomRight");
        w.WriteAttributeString("xrange", "6629,14971");
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("o", "lock", OfficeVmlNs);
        w.WriteAttributeString("v", "ext", VmlNs, "edit");
        w.WriteAttributeString("text", "t");
        w.WriteAttributeString("shapetype", "t");
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static string Points(double points) => points.ToString("0.##", CultureInfo.InvariantCulture) + "pt";
}
