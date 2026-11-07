#nullable enable
namespace Divit.Docx;

using Divit.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

/// <content>
/// The main part and the headers and footers: sections, paragraphs, runs, fields and the table of contents.
/// </content>
public sealed partial class DocxDocument
{
    private const string TocHeadingStyle = "TOCHeading";

    /// <summary>
    /// The main part: the sections one after another, each closed by its properties (the last one by the body's).
    /// </summary>
    private void WriteDocument(XmlWriter w, Parts parts)
    {
        var scope = new Scope(parts, Headings(), Captions());
        w.WriteStartElement("w", "document", Main);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        w.WriteStartElement("body", Main);
        for (var i = 0; i < _sections.Count; i++)
        {
            var section = _sections[i];
            scope.Width = section.TextWidthTwips;
            WriteBlocks(w, section, scope);
            if (i < _sections.Count - 1)
            {
                // A section ends with a paragraph that carries its properties.
                w.WriteStartElement("p", Main);
                w.WriteStartElement("pPr", Main);
                WriteSectionProperties(w, section, parts);
                w.WriteEndElement();
                w.WriteEndElement();
            }
        }

        WriteSectionProperties(w, Section, parts);
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private void WriteHeaderFooter(XmlWriter w, DocxHeaderFooter part, Parts parts)
    {
        w.WriteStartElement("w", part.IsFooter ? "ftr" : "hdr", Main);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        WriteBlocks(w, part, new Scope(parts, Array.Empty<(int, string, string)>(), Array.Empty<(DocxCaptionKind, string)>()) { Width = parts.Widths[part] });

        // A header or footer holds at least one paragraph.
        if (part.Blocks.Count == 0)
        {
            w.WriteElementString("p", Main, null);
        }

        w.WriteEndElement();
    }

    private void WriteBlocks(XmlWriter w, DocxContainer container, Scope scope, DocxAlignment? alignment = null)
    {
        foreach (var block in container.Blocks)
        {
            switch (block)
            {
                case DocxParagraph paragraph:
                    WriteParagraph(w, paragraph, scope, alignment);
                    break;
                case DocxTableOfContents toc:
                    WriteContents(w, toc, scope);
                    break;
                case DocxTable table:
                    WriteTable(w, table, scope);
                    break;
            }
        }
    }

    private void WriteParagraph(XmlWriter w, DocxParagraph p, Scope scope, DocxAlignment? alignment = null)
    {
        w.WriteStartElement("p", Main);
        WriteParagraphProperties(w, p, alignment);
        foreach (var run in p.Runs)
        {
            WriteRun(w, run, scope);
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Paragraph properties in the order of the standard: style, keeping, page break, border, shading, tabs, spacing,
    /// indents, alignment.
    /// </summary>
    private static void WriteParagraphProperties(XmlWriter w, DocxParagraph p, DocxAlignment? fallback)
    {
        var any = p.Style != DocxStyle.Normal || p.KeepWithNext || p.KeepTogether || p.PageBreakBefore || p.Border != null
            || p.Shading != null || p.TabStops.Count > 0 || p.SpaceBefore != null || p.SpaceAfter != null || p.LineSpacing != null
            || p.LeftIndent != 0 || p.RightIndent != 0 || p.FirstLineIndent != 0 || (p.Alignment ?? fallback) != null;
        if (!any)
        {
            return;
        }

        w.WriteStartElement("pPr", Main);
        if (p.Style != DocxStyle.Normal) { Val(w, "pStyle", p.Style); }
        if (p.KeepWithNext) { w.WriteElementString("keepNext", Main, null); }
        if (p.KeepTogether) { w.WriteElementString("keepLines", Main, null); }
        if (p.PageBreakBefore) { w.WriteElementString("pageBreakBefore", Main, null); }
        if (p.Border is DocxBorder border)
        {
            w.WriteStartElement("pBdr", Main);
            foreach (var side in new[] { "top", "left", "bottom", "right" })
            {
                Border(w, side, border, 4);
            }

            w.WriteEndElement();
        }

        if (p.Shading != null) { Shading(w, p.Shading); }
        if (p.TabStops.Count > 0)
        {
            w.WriteStartElement("tabs", Main);
            foreach (var tab in p.TabStops)
            {
                w.WriteStartElement("tab", Main);
                w.WriteAttributeString("w", "val", Main, tab.Alignment == DocxTabAlignment.Right ? "right" : tab.Alignment == DocxTabAlignment.Center ? "center" : "left");
                if (tab.Leader != DocxTabLeader.None)
                {
                    w.WriteAttributeString("w", "leader", Main, tab.Leader == DocxTabLeader.Dot ? "dot" : tab.Leader == DocxTabLeader.Hyphen ? "hyphen" : "underscore");
                }

                w.WriteAttributeString("w", "pos", Main, Int(DocxUnits.Twips(tab.Position)));
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        if (p.SpaceBefore != null || p.SpaceAfter != null || p.LineSpacing != null)
        {
            w.WriteStartElement("spacing", Main);
            if (p.SpaceBefore is double before) { w.WriteAttributeString("w", "before", Main, Int(DocxUnits.PointTwips(before))); }
            if (p.SpaceAfter is double after) { w.WriteAttributeString("w", "after", Main, Int(DocxUnits.PointTwips(after))); }
            if (p.LineSpacing is double line)
            {
                w.WriteAttributeString("w", "line", Main, Int((int)Math.Round(line * 240)));
                w.WriteAttributeString("w", "lineRule", Main, "auto");
            }

            w.WriteEndElement();
        }

        if (p.LeftIndent != 0 || p.RightIndent != 0 || p.FirstLineIndent != 0)
        {
            w.WriteStartElement("ind", Main);
            if (p.LeftIndent != 0) { w.WriteAttributeString("w", "left", Main, Int(DocxUnits.Twips(p.LeftIndent))); }
            if (p.RightIndent != 0) { w.WriteAttributeString("w", "right", Main, Int(DocxUnits.Twips(p.RightIndent))); }
            if (p.FirstLineIndent > 0) { w.WriteAttributeString("w", "firstLine", Main, Int(DocxUnits.Twips(p.FirstLineIndent))); }
            if (p.FirstLineIndent < 0) { w.WriteAttributeString("w", "hanging", Main, Int(DocxUnits.Twips(-p.FirstLineIndent))); }
            w.WriteEndElement();
        }

        if ((p.Alignment ?? fallback) is DocxAlignment alignment)
        {
            Val(w, "jc", alignment == DocxAlignment.Center ? "center" : alignment == DocxAlignment.Right ? "right" : alignment == DocxAlignment.Justify ? "both" : "left");
        }

        w.WriteEndElement();
    }

    private void WriteRun(XmlWriter w, DocxRun run, Scope scope)
    {
        if (run.Field != null || run.Code != null)
        {
            // A simple field: Word computes it; the text inside is shown until then.
            w.WriteStartElement("fldSimple", Main);
            w.WriteAttributeString("w", "instr", Main, run.Code ?? FieldCode(run.Field!.Value));
            w.WriteStartElement("r", Main);
            WriteRunProperties(w, run.Font);
            Text(w, run.Result ?? FieldResult(run.Field!.Value));
            w.WriteEndElement();
            w.WriteEndElement();
            return;
        }

        if (run.Image != null)
        {
            WriteImage(w, run, scope);
            return;
        }

        w.WriteStartElement("r", Main);
        WriteRunProperties(w, run.Font);
        if (run.Text != null) { Text(w, run.Text); }
        if (run.Tab) { w.WriteElementString("tab", Main, null); }
        if (run.LineBreak) { w.WriteElementString("br", Main, null); }
        if (run.PageBreak)
        {
            w.WriteStartElement("br", Main);
            w.WriteAttributeString("w", "type", Main, "page");
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Run properties in the order of the standard; nothing when the run takes its style's look.
    /// </summary>
    private static void WriteRunProperties(XmlWriter w, DocxFont? font)
    {
        if (!(font is DocxFont f))
        {
            return;
        }

        w.WriteStartElement("rPr", Main);
        if (f.Name != null) { Fonts(w, f.Name); }
        if (f.Bold) { w.WriteElementString("b", Main, null); w.WriteElementString("bCs", Main, null); }
        if (f.Italic) { w.WriteElementString("i", Main, null); w.WriteElementString("iCs", Main, null); }
        if (f.AllCaps) { w.WriteElementString("caps", Main, null); }
        if (f.SmallCaps) { w.WriteElementString("smallCaps", Main, null); }
        if (f.Strikethrough) { w.WriteElementString("strike", Main, null); }
        if (f.Color != null) { Val(w, "color", OfficeText.Rgb(f.Color)); }
        if (f.Size > 0) { Size(w, f.Size); }
        if (f.Underline) { Val(w, "u", "single"); }
        if (f.Background != null) { Shading(w, f.Background); }
        if (f.Script != DocxScript.None) { Val(w, "vertAlign", f.Script == DocxScript.Superscript ? "superscript" : "subscript"); }
        w.WriteEndElement();
    }

    private static string FieldCode(DocxField field)
    {
        switch (field)
        {
            case DocxField.NumPages: return " NUMPAGES ";
            case DocxField.SectionPages: return " SECTIONPAGES ";
            case DocxField.Date: return " DATE \\@ \"dd.MM.yyyy\" ";
            case DocxField.Title: return " TITLE ";
            default: return " PAGE ";
        }
    }

    /// <summary>
    /// What a field shows before Word computes it: a page number of 1, the creation date, the title.
    /// </summary>
    private string FieldResult(DocxField field)
    {
        switch (field)
        {
            case DocxField.Date: return Created?.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture) ?? string.Empty;
            case DocxField.Title: return Title ?? string.Empty;
            default: return "1";
        }
    }

    /// <summary>
    /// The table of contents: a TOC field around entries for the headings found now, without page numbers. Word
    /// rebuilds the entries with their pages when it updates the field.
    /// </summary>
    private static void WriteContents(XmlWriter w, DocxTableOfContents toc, Scope scope)
    {
        // Headings by level, or the captions of one kind (a list of figures or tables).
        var entries = new List<(string Style, string Number, string Text)>();
        if (toc.Captions is DocxCaptionKind kind)
        {
            foreach (var caption in scope.Captions)
            {
                if (caption.Kind == kind)
                {
                    entries.Add((FiguresStyle, string.Empty, caption.Text));
                }
            }
        }
        else
        {
            foreach (var heading in scope.Headings)
            {
                if (heading.Level <= toc.Levels)
                {
                    entries.Add(("TOC" + Int(heading.Level), heading.Number, heading.Text));
                }
            }
        }

        if (entries.Count == 0)
        {
            entries.Add((toc.Captions != null ? FiguresStyle : "TOC1", string.Empty, string.Empty));
        }

        var code = toc.Captions is DocxCaptionKind captions
            ? " TOC \\h \\z \\c \"" + SequenceName(captions) + "\" "
            : " TOC \\o \"1-" + Int(toc.Levels) + "\" \\h \\z \\u ";
        for (var i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            w.WriteStartElement("p", Main);
            w.WriteStartElement("pPr", Main);
            Val(w, "pStyle", entry.Style);
            w.WriteEndElement();
            if (i == 0)
            {
                FieldChar(w, "begin");
                w.WriteStartElement("r", Main);
                w.WriteStartElement("instrText", Main);
                w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(code);
                w.WriteEndElement();
                w.WriteEndElement();
                FieldChar(w, "separate");
            }

            var text = entry.Number.Length > 0 ? entry.Number + " " + entry.Text : entry.Text;
            if (text.Length > 0)
            {
                w.WriteStartElement("r", Main);
                Text(w, text);
                w.WriteEndElement();
            }

            if (i == entries.Count - 1)
            {
                FieldChar(w, "end");
            }

            w.WriteEndElement();
        }
    }

    private static void FieldChar(XmlWriter w, string type)
    {
        w.WriteStartElement("r", Main);
        w.WriteStartElement("fldChar", Main);
        w.WriteAttributeString("w", "fldCharType", Main, type);
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The headings of the document in order, with the numbers Word gives them when headings are numbered.
    /// </summary>
    private List<(int Level, string Number, string Text)> Headings()
    {
        var list = new List<(int, string, string)>();
        var counters = new int[7];
        foreach (var section in _sections)
        {
            foreach (var block in section.Blocks)
            {
                if (!(block is DocxParagraph p) || !p.Style.StartsWith("Heading", StringComparison.Ordinal)
                    || p.Style.Length != 8 || p.Style[7] < '1' || p.Style[7] > '6')
                {
                    continue;
                }

                var level = p.Style[7] - '0';
                var number = string.Empty;
                if (NumberHeadings)
                {
                    counters[level]++;
                    for (var deeper = level + 1; deeper < counters.Length; deeper++)
                    {
                        counters[deeper] = 0;
                    }

                    var parts = new string[level];
                    for (var l = 1; l <= level; l++)
                    {
                        parts[l - 1] = Int(counters[l]);
                    }

                    number = string.Join(".", parts);
                }

                list.Add((level, number, p.Text));
            }
        }

        return list;
    }

    /// <summary>
    /// Section properties in the order of the standard: header and footer references, paper, margins, page numbers,
    /// vertical alignment, first page.
    /// </summary>
    private static void WriteSectionProperties(XmlWriter w, DocxSection section, Parts parts)
    {
        w.WriteStartElement("sectPr", Main);
        Reference(w, "headerReference", "default", section.DefaultHeaderPart, parts);
        Reference(w, "headerReference", "even", section.EvenHeaderPart, parts);
        Reference(w, "headerReference", "first", section.FirstHeaderPart, parts);
        Reference(w, "footerReference", "default", section.DefaultFooterPart, parts);
        Reference(w, "footerReference", "even", section.EvenFooterPart, parts);
        Reference(w, "footerReference", "first", section.FirstFooterPart, parts);

        var (width, height) = section.PaperTwips;
        w.WriteStartElement("pgSz", Main);
        w.WriteAttributeString("w", "w", Main, Int(width));
        w.WriteAttributeString("w", "h", Main, Int(height));
        if (section.Landscape) { w.WriteAttributeString("w", "orient", Main, "landscape"); }
        w.WriteEndElement();

        var m = section.Margins;
        w.WriteStartElement("pgMar", Main);
        w.WriteAttributeString("w", "top", Main, Int(DocxUnits.Twips(m.Top)));
        w.WriteAttributeString("w", "right", Main, Int(DocxUnits.Twips(m.Right)));
        w.WriteAttributeString("w", "bottom", Main, Int(DocxUnits.Twips(m.Bottom)));
        w.WriteAttributeString("w", "left", Main, Int(DocxUnits.Twips(m.Left)));
        w.WriteAttributeString("w", "header", Main, Int(DocxUnits.Twips(m.Header)));
        w.WriteAttributeString("w", "footer", Main, Int(DocxUnits.Twips(m.Footer)));
        w.WriteAttributeString("w", "gutter", Main, "0");
        w.WriteEndElement();

        if (section.PageNumberStart != null || section.PageNumberFormat != DocxNumberFormat.Decimal)
        {
            w.WriteStartElement("pgNumType", Main);
            if (section.PageNumberFormat != DocxNumberFormat.Decimal)
            {
                w.WriteAttributeString("w", "fmt", Main, section.PageNumberFormat == DocxNumberFormat.LowerRoman ? "lowerRoman" : "upperRoman");
            }

            if (section.PageNumberStart is int start) { w.WriteAttributeString("w", "start", Main, Int(start)); }
            w.WriteEndElement();
        }

        if (section.VerticalAlignment != DocxVerticalAlignment.Top)
        {
            Val(w, "vAlign", section.VerticalAlignment == DocxVerticalAlignment.Center ? "center" : "bottom");
        }

        if (section.DifferentFirstPage) { w.WriteElementString("titlePg", Main, null); }
        w.WriteEndElement();
    }

    private static void Reference(XmlWriter w, string name, string type, DocxHeaderFooter? part, Parts parts)
    {
        if (part == null)
        {
            return;
        }

        w.WriteStartElement(name, Main);
        w.WriteAttributeString("w", "type", Main, type);
        w.WriteAttributeString("r", "id", RelNs, parts.Ids[part]);
        w.WriteEndElement();
    }

    private static void Text(XmlWriter w, string text)
    {
        w.WriteStartElement("t", Main);
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
        {
            w.WriteAttributeString("xml", "space", null, "preserve");
        }

        w.WriteString(OfficeText.Clean(text));
        w.WriteEndElement();
    }

    private static void Fonts(XmlWriter w, string name)
    {
        w.WriteStartElement("rFonts", Main);
        w.WriteAttributeString("w", "ascii", Main, name);
        w.WriteAttributeString("w", "hAnsi", Main, name);
        w.WriteAttributeString("w", "eastAsia", Main, name);
        w.WriteAttributeString("w", "cs", Main, name);
        w.WriteEndElement();
    }

    private static void Size(XmlWriter w, double points)
    {
        Val(w, "sz", Int(DocxUnits.HalfPoints(points)));
        Val(w, "szCs", Int(DocxUnits.HalfPoints(points)));
    }

    private static void Shading(XmlWriter w, string color)
    {
        w.WriteStartElement("shd", Main);
        w.WriteAttributeString("w", "val", Main, "clear");
        w.WriteAttributeString("w", "color", Main, "auto");
        w.WriteAttributeString("w", "fill", Main, OfficeText.Rgb(color));
        w.WriteEndElement();
    }

    /// <summary>
    /// One border line; <paramref name="space"/> is the gap to the text in points.
    /// </summary>
    private static void Border(XmlWriter w, string side, DocxBorder border, int space)
    {
        w.WriteStartElement(side, Main);
        w.WriteAttributeString("w", "val", Main, border.Style == DocxLineStyle.Double ? "double" : border.Style == DocxLineStyle.Dotted ? "dotted" : border.Style == DocxLineStyle.Dashed ? "dashed" : "single");
        w.WriteAttributeString("w", "sz", Main, Int(DocxUnits.Eighths(border.Width)));
        w.WriteAttributeString("w", "space", Main, Int(space));
        w.WriteAttributeString("w", "color", Main, OfficeText.Rgb(border.Color));
        w.WriteEndElement();
    }

    private static void Val(XmlWriter w, string name, string value)
    {
        w.WriteStartElement(name, Main);
        w.WriteAttributeString("w", "val", Main, value);
        w.WriteEndElement();
    }

    private static string Int(int value) => OfficeText.Int(value);
}
