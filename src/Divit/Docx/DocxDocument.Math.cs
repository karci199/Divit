#nullable enable
namespace Divit.Docx;

using System;
using System.Globalization;
using System.Xml;

/// <content>
/// Equations (Office Math) and tracked changes.
/// </content>
public sealed partial class DocxDocument
{
    private const string MathNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>
    /// An equation: in the line (<c>m:oMath</c>), or set apart and centred (<c>m:oMathPara</c>).
    /// </summary>
    private static void WriteEquation(XmlWriter w, DocxRun run)
    {
        if (run.MathDisplay) { w.WriteStartElement("m", "oMathPara", MathNs); }
        w.WriteStartElement("m", "oMath", MathNs);
        WriteMath(w, run.Math!, run.Font);
        w.WriteEndElement();
        if (run.MathDisplay) { w.WriteEndElement(); }
    }

    /// <summary>
    /// One piece of an equation and its parts, each in the element order of the standard.
    /// </summary>
    private static void WriteMath(XmlWriter w, DocxMath math, DocxFont? font)
    {
        var parts = math.Parts;
        switch (math.Kind)
        {
            case DocxMathKind.Text:
            case DocxMathKind.Plain:
                w.WriteStartElement("m", "r", MathNs);
                if (math.Kind == DocxMathKind.Plain)
                {
                    w.WriteStartElement("m", "rPr", MathNs);
                    MathVal(w, "sty", "p");
                    w.WriteEndElement();
                }

                WriteRunProperties(w, font);
                w.WriteStartElement("m", "t", MathNs);
                var text = math.Content!;
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
                {
                    w.WriteAttributeString("xml", "space", null, "preserve");
                }

                w.WriteString(Divit.Packaging.OfficeText.Clean(text));
                w.WriteEndElement();
                w.WriteEndElement();
                break;
            case DocxMathKind.Row:
                foreach (var part in parts)
                {
                    WriteMath(w, part, font);
                }

                break;
            case DocxMathKind.Fraction:
                w.WriteStartElement("m", "f", MathNs);
                MathArgument(w, "num", parts[0], font);
                MathArgument(w, "den", parts[1], font);
                w.WriteEndElement();
                break;
            case DocxMathKind.Sup:
                w.WriteStartElement("m", "sSup", MathNs);
                MathArgument(w, "e", parts[0], font);
                MathArgument(w, "sup", parts[1], font);
                w.WriteEndElement();
                break;
            case DocxMathKind.Sub:
                w.WriteStartElement("m", "sSub", MathNs);
                MathArgument(w, "e", parts[0], font);
                MathArgument(w, "sub", parts[1], font);
                w.WriteEndElement();
                break;
            case DocxMathKind.SubSup:
                w.WriteStartElement("m", "sSubSup", MathNs);
                MathArgument(w, "e", parts[0], font);
                MathArgument(w, "sub", parts[1], font);
                MathArgument(w, "sup", parts[2], font);
                w.WriteEndElement();
                break;
            case DocxMathKind.Root:
                w.WriteStartElement("m", "rad", MathNs);
                if (parts.Count == 1)
                {
                    w.WriteStartElement("m", "radPr", MathNs);
                    MathVal(w, "degHide", "1");
                    w.WriteEndElement();
                }

                MathArgument(w, "deg", parts.Count > 1 ? parts[1] : null, font);
                MathArgument(w, "e", parts[0], font);
                w.WriteEndElement();
                break;
            case DocxMathKind.Brackets:
                w.WriteStartElement("m", "d", MathNs);
                if (math.Content != "(" || math.Close != ")")
                {
                    w.WriteStartElement("m", "dPr", MathNs);
                    MathVal(w, "begChr", math.Content!);
                    MathVal(w, "endChr", math.Close!);
                    w.WriteEndElement();
                }

                MathArgument(w, "e", parts[0], font);
                w.WriteEndElement();
                break;
            case DocxMathKind.Operator:
                w.WriteStartElement("m", "nary", MathNs);
                w.WriteStartElement("m", "naryPr", MathNs);
                MathVal(w, "chr", math.Content!);
                var integral = math.Content!.Length == 1 && math.Content[0] >= '∫' && math.Content[0] <= '∳';
                MathVal(w, "limLoc", integral ? "subSup" : "undOvr");
                if (!math.HasLower) { MathVal(w, "subHide", "1"); }
                if (!math.HasUpper) { MathVal(w, "supHide", "1"); }
                w.WriteEndElement();
                MathArgument(w, "sub", math.HasLower ? parts[1] : null, font);
                MathArgument(w, "sup", math.HasUpper ? parts[2] : null, font);
                MathArgument(w, "e", parts[0], font);
                w.WriteEndElement();
                break;
            case DocxMathKind.Function:
                w.WriteStartElement("m", "func", MathNs);
                MathArgument(w, "fName", parts[0], font);
                MathArgument(w, "e", parts[1], font);
                w.WriteEndElement();
                break;
        }
    }

    /// <summary>
    /// An argument of a math element (numerator, base, index …); empty when there is none.
    /// </summary>
    private static void MathArgument(XmlWriter w, string name, DocxMath? content, DocxFont? font)
    {
        w.WriteStartElement("m", name, MathNs);
        if (content != null) { WriteMath(w, content, font); }
        w.WriteEndElement();
    }

    private static void MathVal(XmlWriter w, string name, string value)
    {
        w.WriteStartElement("m", name, MathNs);
        w.WriteAttributeString("m", "val", MathNs, value);
        w.WriteEndElement();
    }

    /// <summary>
    /// A tracked insertion or deletion: the run inside a revision mark with its author and time; deleted text is
    /// <c>delText</c>.
    /// </summary>
    private void WriteRevision(XmlWriter w, DocxRun run, Parts parts)
    {
        var deletion = run.Revision == DocxRevision.Deletion;
        w.WriteStartElement(deletion ? "del" : "ins", Main);
        w.WriteAttributeString("w", "id", Main, Int(parts.Annotations++));
        w.WriteAttributeString("w", "author", Main, run.RevisionAuthor ?? Author ?? Application);
        if ((run.RevisionDate ?? Created) is DateTime date)
        {
            w.WriteAttributeString("w", "date", Main, date.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));
        }

        w.WriteStartElement("r", Main);
        WriteRunProperties(w, run.Font);
        var text = run.Text ?? string.Empty;
        w.WriteStartElement(deletion ? "delText" : "t", Main);
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
        {
            w.WriteAttributeString("xml", "space", null, "preserve");
        }

        w.WriteString(Divit.Packaging.OfficeText.Clean(text));
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }
}
