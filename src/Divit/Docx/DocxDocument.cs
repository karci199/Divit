#nullable enable
namespace Divit.Docx;

using Divit.Packaging;
using System;
using System.Collections.Generic;

/// <summary>
/// A Word document (.docx): sections of paragraphs, headings, tables, headers and footers, written to disk with
/// <see cref="Save"/>.
/// </summary>
/// <remarks>
/// <para>Content goes into the last section; <see cref="AddSection"/> starts a new one on a new page, with its own
/// page size, margins, headers and footers. The look of body text, headings and tables comes from
/// <see cref="Look"/>, turned into Word's styles when the file is saved, so a reader can still restyle the document
/// in Word.</para>
/// <para>Word lays out the pages, so page numbers and the table of contents are fields that Word fills in; with a
/// table of contents, Word offers to update the fields when it opens the file.</para>
/// </remarks>
public sealed partial class DocxDocument
{
    private const string Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.";

    private readonly List<DocxSection> _sections = new List<DocxSection>();

    /// <summary>
    /// Captions so far, per kind: the number the next one gets.
    /// </summary>
    private int _figures;

    private int _tables;

    /// <summary>
    /// Creates an empty document with one A4 section.
    /// </summary>
    public DocxDocument()
    {
        _sections.Add(new DocxSection(null));
    }

    /// <summary>
    /// The house style: fonts, sizes and colours of text, headings and tables.
    /// </summary>
    public DocxLook Look { get; } = new DocxLook();

    /// <summary>
    /// The theme: palette and fonts of Word's pickers. Its fonts default to the look's fonts.
    /// </summary>
    public OfficeTheme Theme { get; } = new OfficeTheme();

    /// <summary>
    /// Numbers the headings by level: 1, 1.1, 1.1.1 …, as a numbered outline Word keeps up to date.
    /// </summary>
    public bool NumberHeadings { get; set; }

    /// <summary>
    /// Language of the text, for spelling and hyphenation, for example <c>tr-TR</c>; Word's own when empty.
    /// </summary>
    public string? Language { get; set; }

    /// <summary>
    /// File property: title. <see cref="DocxField.Title"/> shows it in the text.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// File property: subject.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// File property: keywords.
    /// </summary>
    public string? Keywords { get; set; }

    /// <summary>
    /// File property: description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// File property: author; <see cref="Application"/> when empty.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Application name written into the file properties.
    /// </summary>
    public string Application { get; set; } = "Divit";

    /// <summary>
    /// File property: creation time (UTC); the time of saving when empty. Set it to get byte-identical files from
    /// identical input.
    /// </summary>
    public DateTime? Created { get; set; }

    /// <summary>
    /// The sections in order.
    /// </summary>
    public IReadOnlyList<DocxSection> Sections => _sections;

    /// <summary>
    /// The last section, which new content goes into.
    /// </summary>
    public DocxSection Section => _sections[_sections.Count - 1];

    /// <summary>
    /// Starts a new section on a new page; it takes the page settings of the previous one, and its headers and
    /// footers until it gets its own.
    /// </summary>
    /// <returns>The new section.</returns>
    public DocxSection AddSection()
    {
        var section = new DocxSection(Section);
        _sections.Add(section);
        return section;
    }

    /// <summary>
    /// Adds a paragraph to the last section.
    /// </summary>
    /// <param name="text">Its text; none when empty.</param>
    /// <param name="font">Font of the text; the style's when empty.</param>
    /// <param name="style">Paragraph style; <see cref="DocxStyle.Normal"/> when empty.</param>
    /// <returns>The paragraph.</returns>
    public DocxParagraph AddParagraph(string? text = null, DocxFont? font = null, string? style = null) =>
        Section.AddParagraph(text, font, style);

    /// <summary>
    /// Adds a heading; it shows in the table of contents and, with <see cref="NumberHeadings"/>, gets its number.
    /// </summary>
    /// <param name="text">The heading text.</param>
    /// <param name="level">Level 1–6.</param>
    /// <returns>The paragraph.</returns>
    public DocxParagraph AddHeading(string text, int level = 1) => Section.AddParagraph(text, null, DocxStyle.Heading(level));

    /// <summary>
    /// Adds a title in the large title style, as on a cover page.
    /// </summary>
    /// <param name="text">The title.</param>
    /// <returns>The paragraph.</returns>
    public DocxParagraph AddTitle(string text) => Section.AddParagraph(text, null, DocxStyle.Title);

    /// <summary>
    /// Adds a subtitle under a title.
    /// </summary>
    /// <param name="text">The subtitle.</param>
    /// <returns>The paragraph.</returns>
    public DocxParagraph AddSubtitle(string text) => Section.AddParagraph(text, null, DocxStyle.Subtitle);

    /// <summary>
    /// Continues on a new page within the same section.
    /// </summary>
    public void AddPageBreak() => Section.AddParagraph().AddPageBreakRun();

    /// <summary>
    /// Adds a table of contents listing the headings down to a level, with dot leaders and page numbers.
    /// </summary>
    /// <param name="levels">Deepest heading level listed, 1–6.</param>
    /// <param name="title">Title above the list, not listed itself; none when empty.</param>
    /// <returns>The table of contents.</returns>
    public DocxTableOfContents AddTableOfContents(int levels = 3, string? title = "Contents")
    {
        if (title != null)
        {
            Section.AddParagraph(title, null, TocHeadingStyle);
        }

        var toc = new DocxTableOfContents(Math.Max(1, Math.Min(6, levels)));
        Section.AddBlock(toc);
        return toc;
    }

    /// <summary>
    /// Adds a table to the last section.
    /// </summary>
    /// <param name="columns">Number of columns, 1–63.</param>
    /// <returns>The table.</returns>
    public DocxTable AddTable(int columns) => Section.AddTable(columns);

    /// <summary>
    /// Adds a numbered caption: "Figure 3 – Test set-up". Word keeps the numbers in order and lists the captions in a
    /// list of figures or tables (<see cref="AddTableOfFigures"/>). Put a table's caption above it, a figure's below.
    /// </summary>
    /// <param name="kind">What the caption numbers.</param>
    /// <param name="text">The caption text.</param>
    /// <returns>The caption paragraph.</returns>
    public DocxParagraph AddCaption(DocxCaptionKind kind, string text)
    {
        var number = kind == DocxCaptionKind.Figure ? ++_figures : ++_tables;
        var label = kind == DocxCaptionKind.Figure ? Look.FigureLabel : Look.TableLabel;
        var paragraph = Section.AddParagraph(label + " ", null, DocxStyle.Caption);
        paragraph.AddFieldRun(" SEQ " + SequenceName(kind) + " \\* ARABIC ", Int(number));
        paragraph.Add(Look.CaptionSeparator + text);
        paragraph.Caption = kind;
        paragraph.CaptionText = label + " " + Int(number) + Look.CaptionSeparator + text;
        paragraph.KeepWithNext = kind == DocxCaptionKind.Table;
        return paragraph;
    }

    /// <summary>
    /// Adds a list of the figures or the tables with their pages, like a table of contents of the captions.
    /// </summary>
    /// <param name="kind">What it lists.</param>
    /// <param name="title">Title above the list, not listed itself; none when empty.</param>
    /// <returns>The list.</returns>
    public DocxTableOfContents AddTableOfFigures(DocxCaptionKind kind, string? title = null)
    {
        if (title != null)
        {
            Section.AddParagraph(title, null, TocHeadingStyle);
        }

        var list = new DocxTableOfContents(1) { Captions = kind };
        Section.AddBlock(list);
        return list;
    }

    /// <summary>
    /// The field name that counts a caption kind; Word's own, so its Insert Caption continues the numbering.
    /// </summary>
    private static string SequenceName(DocxCaptionKind kind) => kind == DocxCaptionKind.Figure ? "Figure" : "Table";

    /// <summary>
    /// Writes the document; an existing target is replaced only once the new file is complete.
    /// </summary>
    /// <param name="path">Target file.</param>
    public void Save(string path)
    {
        var parts = PlanParts();
        OpcPackage.Save(path, Created, package =>
        {
            package.XmlPart("word/document.xml", w => WriteDocument(w, parts));
            foreach (var part in parts.HeadersAndFooters)
            {
                var name = (part.IsFooter ? "footer" : "header") + part.Number.ToString(System.Globalization.CultureInfo.InvariantCulture);
                package.XmlPart("word/" + name + ".xml", w => WriteHeaderFooter(w, part, parts));
                package.Relationships("word/_rels/" + name + ".xml.rels", parts.PartRelationships[part]);
            }

            foreach (var media in parts.Media)
            {
                package.BinaryPart("word/media/" + media.Name, media.Data, false);
            }

            package.XmlPart("word/styles.xml", WriteStyles);
            if (NumberHeadings)
            {
                package.XmlPart("word/numbering.xml", WriteNumbering);
            }

            package.XmlPart("word/settings.xml", w => WriteSettings(w, parts));
            package.XmlPart("word/theme/theme1.xml", w => ThemeWriter.Write(w, Theme, Look.BodyFont));
            package.Relationships("word/_rels/document.xml.rels", parts.Relationships);
            package.DocumentProperties(new DocumentProperties(Application)
            {
                Title = Title,
                Subject = Subject,
                Keywords = Keywords,
                Description = Description,
                Author = Author,
                Created = Created,
            });

            var root = new RelationshipList();
            root.Add(RelationshipTypes.OfficeDocument, "word/document.xml");
            root.Add(RelationshipTypes.CoreProperties, "docProps/core.xml");
            root.Add(RelationshipTypes.ExtendedProperties, "docProps/app.xml");
            package.Relationships("_rels/.rels", root);

            package.Default("rels", ContentTypes.Relationships);
            package.Default("xml", ContentTypes.Xml);
            foreach (var type in parts.MediaTypes)
            {
                package.Default(type.Key, type.Value);
            }

            package.Override("/word/document.xml", ContentType + "document.main+xml");
            package.Override("/word/styles.xml", ContentType + "styles+xml");
            package.Override("/word/settings.xml", ContentType + "settings+xml");
            if (NumberHeadings)
            {
                package.Override("/word/numbering.xml", ContentType + "numbering+xml");
            }

            foreach (var part in parts.HeadersAndFooters)
            {
                var name = (part.IsFooter ? "footer" : "header") + part.Number.ToString(System.Globalization.CultureInfo.InvariantCulture);
                package.Override("/word/" + name + ".xml", ContentType + (part.IsFooter ? "footer+xml" : "header+xml"));
            }

            package.Override("/word/theme/theme1.xml", ThemeWriter.ContentType);
            package.Override("/docProps/core.xml", ContentTypes.CoreProperties);
            package.Override("/docProps/app.xml", ContentTypes.ExtendedProperties);
        });
    }

    /// <summary>
    /// Numbers the headers and footers and gives every part its relationship id before anything is written.
    /// </summary>
    private Parts PlanParts()
    {
        var parts = new Parts();
        parts.Relationships.Add(RelationshipTypes.Styles, "styles.xml");
        parts.Relationships.Add(RelType + "settings", "settings.xml");
        parts.Relationships.Add(RelationshipTypes.Theme, "theme/theme1.xml");
        if (NumberHeadings)
        {
            parts.Relationships.Add(RelType + "numbering", "numbering.xml");
        }

        int headers = 0, footers = 0;
        foreach (var section in _sections)
        {
            foreach (var part in new[] { section.DefaultHeaderPart, section.FirstHeaderPart, section.EvenHeaderPart, section.DefaultFooterPart, section.FirstFooterPart, section.EvenFooterPart })
            {
                if (part == null)
                {
                    continue;
                }

                part.Number = part.IsFooter ? ++footers : ++headers;
                parts.HeadersAndFooters.Add(part);
                var name = (part.IsFooter ? "footer" : "header") + part.Number.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".xml";
                parts.Ids[part] = parts.Relationships.Add(RelType + (part.IsFooter ? "footer" : "header"), name);
                parts.Widths[part] = section.TextWidthTwips;
                var own = new RelationshipList();
                parts.PartRelationships[part] = own;
                PlanImages(part, own, parts);
            }

            parts.EvenPages |= section.EvenHeaderPart != null || section.EvenFooterPart != null;
        }

        foreach (var section in _sections)
        {
            PlanImages(section, parts.Relationships, parts);
            foreach (var block in section.Blocks)
            {
                parts.HasContents |= block is DocxTableOfContents;
            }
        }

        return parts;
    }

    /// <summary>
    /// Gives every image of a part (in its paragraphs and table cells, nested too) a media file and a relationship
    /// of that part; images are stored as they are.
    /// </summary>
    private static void PlanImages(DocxContainer container, RelationshipList relationships, Parts parts)
    {
        foreach (var block in container.Blocks)
        {
            if (block is DocxParagraph paragraph)
            {
                foreach (var run in paragraph.Runs)
                {
                    if (run.Image == null)
                    {
                        continue;
                    }

                    var name = "image" + (parts.Media.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "." + run.ImageInfo.Extension;
                    parts.Media.Add((name, run.Image));
                    parts.MediaTypes[run.ImageInfo.Extension] = run.ImageInfo.ContentType;
                    parts.ImageIds[run] = relationships.Add(RelationshipTypes.Image, "media/" + name);
                }
            }
            else if (block is DocxTable table)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        PlanImages(cell, relationships, parts);
                    }
                }
            }
        }
    }

    /// <summary>
    /// The parts of one save: relationships of the main part, the headers and footers with their ids.
    /// </summary>
    private sealed class Parts
    {
        public RelationshipList Relationships { get; } = new RelationshipList();

        public List<DocxHeaderFooter> HeadersAndFooters { get; } = new List<DocxHeaderFooter>();

        public Dictionary<DocxHeaderFooter, string> Ids { get; } = new Dictionary<DocxHeaderFooter, string>();

        /// <summary>
        /// Width between the margins of the section a header or footer belongs to.
        /// </summary>
        public Dictionary<DocxHeaderFooter, int> Widths { get; } = new Dictionary<DocxHeaderFooter, int>();

        /// <summary>
        /// The relationships of each header and footer part (its images).
        /// </summary>
        public Dictionary<DocxHeaderFooter, RelationshipList> PartRelationships { get; } = new Dictionary<DocxHeaderFooter, RelationshipList>();

        /// <summary>
        /// Media files in order: name in word/media and bytes.
        /// </summary>
        public List<(string Name, byte[] Data)> Media { get; } = new List<(string, byte[])>();

        /// <summary>
        /// Content types of the image formats used, by extension, in order of first use.
        /// </summary>
        public SortedDictionary<string, string> MediaTypes { get; } = new SortedDictionary<string, string>(StringComparer.Ordinal);

        /// <summary>
        /// The relationship id of each image, in the part that holds it.
        /// </summary>
        public Dictionary<DocxRun, string> ImageIds { get; } = new Dictionary<DocxRun, string>();

        /// <summary>
        /// Counter of drawing ids, which Word wants unique in the document.
        /// </summary>
        public int Drawings { get; set; }

        /// <summary>
        /// Whether any section has its own even-page header or footer, which Word turns on for the whole document.
        /// </summary>
        public bool EvenPages { get; set; }

        /// <summary>
        /// Whether the document has a table of contents, which asks Word to update the fields on opening.
        /// </summary>
        public bool HasContents { get; set; }
    }
}
