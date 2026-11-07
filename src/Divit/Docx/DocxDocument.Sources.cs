#nullable enable
namespace Divit.Docx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

/// <content>
/// Sources, citations and the bibliography, kept where Word's source manager reads them: a custom XML part.
/// </content>
public sealed partial class DocxDocument
{
    private const string BibliographyNs = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
    private const string CustomXmlNs = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
    private const string BibliographyStyleId = "Bibliography";

    /// <summary>
    /// The id of the custom XML part with the sources; fixed, since a document has one.
    /// </summary>
    private const string SourcesItemId = "{6F2A3C4E-9B1D-4E5A-8C7F-0D1E2F3A4B5C}";

    /// <summary>
    /// A citation: a CITATION field Word formats in the bibliography style; it shows Divit's rendering until then.
    /// </summary>
    private void WriteCitation(XmlWriter w, DocxRun run)
    {
        var source = run.Citation!;
        var code = " CITATION " + source.Tag + " \\l " + Int(LanguageId()) + (run.CitedPages != null ? " \\p " + run.CitedPages : string.Empty) + " ";
        w.WriteStartElement("fldSimple", Main);
        w.WriteAttributeString("w", "instr", Main, code);
        w.WriteStartElement("r", Main);
        WriteRunProperties(w, run.Font);
        Text(w, CitationText(source, run.CitedPages));
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The bibliography: a BIBLIOGRAPHY field around one paragraph per source, which Word writes again in its
    /// style when it updates the field.
    /// </summary>
    private void WriteBibliography(XmlWriter w)
    {
        var count = Math.Max(1, _sources.Count);
        for (var i = 0; i < count; i++)
        {
            w.WriteStartElement("p", Main);
            w.WriteStartElement("pPr", Main);
            Val(w, "pStyle", BibliographyStyleId);
            w.WriteEndElement();
            if (i == 0)
            {
                FieldChar(w, "begin");
                w.WriteStartElement("r", Main);
                w.WriteStartElement("instrText", Main);
                w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(" BIBLIOGRAPHY ");
                w.WriteEndElement();
                w.WriteEndElement();
                FieldChar(w, "separate");
            }

            if (i < _sources.Count)
            {
                // The number, a tab to the hanging indent, the entry.
                var entry = EntryText(_sources[i]).Split('\t');
                for (var piece = 0; piece < entry.Length; piece++)
                {
                    w.WriteStartElement("r", Main);
                    if (piece > 0) { w.WriteElementString("tab", Main, null); }
                    Text(w, entry[piece]);
                    w.WriteEndElement();
                }
            }

            if (i == count - 1)
            {
                FieldChar(w, "end");
            }

            w.WriteEndElement();
        }
    }

    /// <summary>
    /// The sources part, in the bibliography schema with the chosen style.
    /// </summary>
    private void WriteSources(XmlWriter w)
    {
        var ieee = BibliographyStyle == DocxBibliographyStyle.Ieee;
        w.WriteStartElement("b", "Sources", BibliographyNs);
        w.WriteAttributeString("SelectedStyle", ieee ? "\\IEEE2006OfficeOnline.xsl" : "\\APASixthEditionOfficeOnline.xsl");
        w.WriteAttributeString("StyleName", ieee ? "IEEE" : "APA");
        w.WriteAttributeString("Version", ieee ? "2006" : "6");
        foreach (var source in _sources)
        {
            w.WriteStartElement("b", "Source", BibliographyNs);
            w.WriteElementString("b", "Tag", BibliographyNs, source.Tag);
            w.WriteElementString("b", "SourceType", BibliographyNs, source.Kind.ToString());
            if (source.Corporate != null || source.Authors.Count > 0)
            {
                w.WriteStartElement("b", "Author", BibliographyNs);
                w.WriteStartElement("b", "Author", BibliographyNs);
                if (source.Authors.Count > 0)
                {
                    w.WriteStartElement("b", "NameList", BibliographyNs);
                    foreach (var (last, first) in source.Authors)
                    {
                        w.WriteStartElement("b", "Person", BibliographyNs);
                        w.WriteElementString("b", "Last", BibliographyNs, Divit.Packaging.OfficeText.Clean(last));
                        if (first.Length > 0) { w.WriteElementString("b", "First", BibliographyNs, Divit.Packaging.OfficeText.Clean(first)); }
                        w.WriteEndElement();
                    }

                    w.WriteEndElement();
                }
                else
                {
                    w.WriteElementString("b", "Corporate", BibliographyNs, Divit.Packaging.OfficeText.Clean(source.Corporate!));
                }

                w.WriteEndElement();
                w.WriteEndElement();
            }

            SourceField(w, "Title", source.Title);
            SourceField(w, "Year", source.Year);
            SourceField(w, "City", source.City);
            SourceField(w, "Publisher", source.Publisher);
            SourceField(w, source.Kind == DocxSourceKind.ConferenceProceedings ? "ConferenceName" : "JournalName", source.Journal);
            SourceField(w, "Volume", source.Volume);
            SourceField(w, "Issue", source.Issue);
            SourceField(w, "Pages", source.Pages);
            SourceField(w, "Edition", source.Edition);
            SourceField(w, "URL", source.Url);
            SourceField(w, "RefOrder", Int(source.Number));
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// The properties of the sources part: its id and schema, by which Word finds it.
    /// </summary>
    private static void WriteSourceProperties(XmlWriter w)
    {
        w.WriteStartElement("ds", "datastoreItem", CustomXmlNs);
        w.WriteAttributeString("ds", "itemID", CustomXmlNs, SourcesItemId);
        w.WriteStartElement("ds", "schemaRefs", CustomXmlNs);
        w.WriteStartElement("ds", "schemaRef", CustomXmlNs);
        w.WriteAttributeString("ds", "uri", CustomXmlNs, BibliographyNs);
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteEndElement();
    }

    private static void SourceField(XmlWriter w, string name, string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            w.WriteElementString("b", name, BibliographyNs, Divit.Packaging.OfficeText.Clean(value!));
        }
    }

    /// <summary>
    /// The Windows language id of the document's language for citation fields; US English when unknown.
    /// </summary>
    private int LanguageId()
    {
        if (Language == null)
        {
            return 1033;
        }

        try
        {
            return CultureInfo.GetCultureInfo(Language).LCID;
        }
        catch (CultureNotFoundException)
        {
            return 1033;
        }
    }

    /// <summary>
    /// A citation as the style writes it: "[1]", or "(Balanis, 2016, p. 12)".
    /// </summary>
    private string CitationText(DocxSource source, string? pages)
    {
        if (BibliographyStyle == DocxBibliographyStyle.Ieee)
        {
            return "[" + Int(source.Number) + (pages != null ? ", p. " + pages : string.Empty) + "]";
        }

        var who = source.Authors.Count == 0 ? source.Corporate ?? source.Title
            : source.Authors.Count == 1 ? source.Authors[0].Last
            : source.Authors.Count == 2 ? source.Authors[0].Last + " & " + source.Authors[1].Last
            : source.Authors[0].Last + " et al.";
        return "(" + who + (source.Year != null ? ", " + source.Year : string.Empty) + (pages != null ? ", p. " + pages : string.Empty) + ")";
    }

    /// <summary>
    /// An entry of the bibliography as the style writes it, close enough to stand until Word formats it.
    /// </summary>
    private string EntryText(DocxSource source)
    {
        var parts = new List<string>();
        var ieee = BibliographyStyle == DocxBibliographyStyle.Ieee;
        var names = new List<string>();
        foreach (var (last, first) in source.Authors)
        {
            names.Add(ieee ? Initials(first) + last : last + (first.Length > 0 ? ", " + Initials(first).TrimEnd() : string.Empty));
        }

        var authors = names.Count == 0 ? source.Corporate
            : names.Count == 1 ? names[0]
            : string.Join(", ", names.GetRange(0, names.Count - 1)) + (ieee ? " and " : ", & ") + names[names.Count - 1];
        if (ieee)
        {
            if (authors != null) { parts.Add(authors); }
            parts.Add("“" + source.Title + ",”");
            if (source.Journal != null) { parts.Add(source.Journal); }
            if (source.Volume != null) { parts.Add("vol. " + source.Volume); }
            if (source.Issue != null) { parts.Add("no. " + source.Issue); }
            if (source.Pages != null) { parts.Add("pp. " + source.Pages); }
            if (source.Edition != null) { parts.Add(source.Edition + " ed."); }
            if (source.City != null) { parts.Add(source.City); }
            if (source.Publisher != null) { parts.Add(source.Publisher); }
            if (source.Year != null) { parts.Add(source.Year); }
            if (source.Url != null) { parts.Add(source.Url); }
            return "[" + Int(source.Number) + "]\t" + string.Join(", ", parts).Replace(",”,", ",”") + ".";
        }

        var text = new StringBuilder();
        text.Append(authors ?? source.Title);
        text.Append(" (").Append(source.Year ?? "n.d.").Append("). ");
        if (authors != null) { text.Append(source.Title).Append(". "); }
        if (source.Journal != null)
        {
            text.Append(source.Journal);
            if (source.Volume != null) { text.Append(", ").Append(source.Volume); }
            if (source.Issue != null) { text.Append('(').Append(source.Issue).Append(')'); }
            if (source.Pages != null) { text.Append(", ").Append(source.Pages); }
            text.Append(". ");
        }

        if (source.City != null) { text.Append(source.City).Append(": "); }
        if (source.Publisher != null) { text.Append(source.Publisher).Append(". "); }
        if (source.Url != null) { text.Append(source.Url); }
        return text.ToString().TrimEnd();
    }

    /// <summary>
    /// Given names as initials with a space after each: "Constantine A." gives "C. A. ".
    /// </summary>
    private static string Initials(string first)
    {
        var text = new StringBuilder();
        foreach (var name in first.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
        {
            text.Append(name[0]).Append(". ");
        }

        return text.ToString();
    }
}
