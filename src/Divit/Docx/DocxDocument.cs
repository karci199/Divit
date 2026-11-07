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

    private int _equations;

    /// <summary>
    /// The sources of the bibliography in order, and their tags.
    /// </summary>
    private readonly List<DocxSource> _sources = new List<DocxSource>();

    private readonly HashSet<string> _sourceTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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
    /// Records every change readers make as a tracked change (Review, Track Changes turned on).
    /// </summary>
    public bool TrackChanges { get; set; }

    /// <summary>
    /// What readers may change once the document is protected; not protected when empty.
    /// </summary>
    public DocxProtection? Protection { get; private set; }

    /// <summary>
    /// Password that lifts the protection; none when empty.
    /// </summary>
    internal string? Password { get; private set; }

    /// <summary>
    /// Protects the document against changes (Review, Restrict Editing). The password is stored as a SHA-512 hash;
    /// this is no encryption, the text stays readable.
    /// </summary>
    /// <param name="protection">What readers may still change.</param>
    /// <param name="password">Password that lifts it; without one, readers can lift it with a click.</param>
    public void Protect(DocxProtection protection = DocxProtection.ReadOnly, string? password = null)
    {
        Protection = protection;
        Password = string.IsNullOrEmpty(password) ? null : password;
    }

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
    /// Starts a bulleted or numbered list in the last section; its items follow where they are added.
    /// </summary>
    /// <param name="kind">Bullets or numbers.</param>
    /// <returns>The list, to add items.</returns>
    public DocxList AddList(DocxListKind kind = DocxListKind.Bullet) => Section.AddList(kind);

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
        var number = kind == DocxCaptionKind.Figure ? ++_figures : kind == DocxCaptionKind.Table ? ++_tables : ++_equations;
        var label = kind == DocxCaptionKind.Figure ? Look.FigureLabel : kind == DocxCaptionKind.Table ? Look.TableLabel : SequenceName(kind);
        var paragraph = Section.AddParagraph(label + " ", null, DocxStyle.Caption);
        paragraph.AddFieldRun(" SEQ " + SequenceName(kind) + " \\* ARABIC ", Int(number));
        paragraph.Add(Look.CaptionSeparator + text);
        paragraph.Caption = kind;
        paragraph.CaptionText = label + " " + Int(number) + Look.CaptionSeparator + text;
        paragraph.CaptionLabel = label + " " + Int(number);
        paragraph.KeepWithNext = kind == DocxCaptionKind.Table;
        return paragraph;
    }

    /// <summary>
    /// Adds an equation set apart in a paragraph of its own: centred, or with a number "(1)" at the right margin
    /// that Word keeps in order and a <see cref="DocxReferenceKind.Label"/> reference can show.
    /// </summary>
    /// <param name="math">The equation.</param>
    /// <param name="numbered">Gives it a number.</param>
    /// <returns>The equation's paragraph; with a number, the number's, which a reference points at.</returns>
    public DocxParagraph AddEquation(DocxMath math, bool numbered = false)
    {
        if (math == null)
        {
            throw new ArgumentNullException(nameof(math));
        }

        if (!numbered)
        {
            var paragraph = Section.AddParagraph();
            paragraph.AddDisplayMath(math);
            return paragraph;
        }

        // A line with the number beside the equation would set it small, as math in running text; a table without
        // lines keeps it at full size, centred, with the number at the right margin.
        var width = Section.TextWidthTwips * 2.54 / 1440;
        var table = Section.AddTable(3).SetColumnWidths(1.5, width - 3, 1.5);
        table.Look = DocxTableLook.Plain;
        var row = table.AddRow();
        row.Cells[1].VerticalAlignment = DocxCellAlignment.Center;
        row.Cells[1].AddParagraph().AddDisplayMath(math);
        row.Cells[2].VerticalAlignment = DocxCellAlignment.Center;
        var label = row.Cells[2].AddParagraph();
        label.Alignment = DocxAlignment.Right;
        var number = Int(++_equations);
        label.Add("(");
        label.AddFieldRun(" SEQ " + SequenceName(DocxCaptionKind.Equation) + " \\* ARABIC ", number);
        label.Add(")");
        label.Caption = DocxCaptionKind.Equation;
        label.CaptionLabel = "(" + number + ")";
        label.CaptionText = label.CaptionLabel;
        label.LabelCount = 3;
        return label;
    }

    /// <summary>
    /// How citations and the bibliography are written.
    /// </summary>
    public DocxBibliographyStyle BibliographyStyle { get; set; }

    /// <summary>
    /// The sources of the bibliography in order.
    /// </summary>
    public IReadOnlyList<DocxSource> Sources => _sources;

    /// <summary>
    /// Adds a source to the document's bibliography, where Word's source manager shows it; cite it with
    /// <see cref="DocxParagraph.AddCitation"/>.
    /// </summary>
    /// <param name="tag">Short name, unique in the document, without spaces, for example <c>Balanis2016</c>.</param>
    /// <param name="kind">The kind of source.</param>
    /// <param name="title">Its title.</param>
    /// <returns>The source, to add authors and details.</returns>
    /// <exception cref="ArgumentException">The tag is empty, has spaces or is taken.</exception>
    public DocxSource AddSource(string tag, DocxSourceKind kind, string title)
    {
        if (string.IsNullOrEmpty(tag) || tag.IndexOfAny(new[] { ' ', '\t', '\\', '"' }) >= 0 || !_sourceTags.Add(tag))
        {
            throw new ArgumentException(Divit.Properties.Messages.Get(Divit.Properties.Strings.SourceTagInvalid), nameof(tag));
        }

        var source = new DocxSource(tag, kind, title ?? string.Empty) { Number = _sources.Count + 1 };
        _sources.Add(source);
        return source;
    }

    /// <summary>
    /// Adds the bibliography: the document's sources, which Word writes out in the bibliography style.
    /// </summary>
    /// <param name="title">Title above the list, not numbered; none when empty.</param>
    /// <returns>The bibliography.</returns>
    public DocxBibliography AddBibliography(string? title = "References")
    {
        if (title != null)
        {
            Section.AddParagraph(title, null, TocHeadingStyle);
        }

        var bibliography = new DocxBibliography();
        Section.AddBlock(bibliography);
        return bibliography;
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
    private static string SequenceName(DocxCaptionKind kind) => kind == DocxCaptionKind.Figure ? "Figure" : kind == DocxCaptionKind.Table ? "Table" : "Equation";

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

            package.XmlPart("word/styles.xml", w => WriteStyles(w, parts));
            if (parts.HasNumbering)
            {
                package.XmlPart("word/numbering.xml", w => WriteNumbering(w, parts));
            }

            if (parts.Footnotes.Count > 0)
            {
                package.XmlPart("word/footnotes.xml", w => WriteFootnotes(w, parts));
            }

            if (_sources.Count > 0)
            {
                package.XmlPart("customXml/item1.xml", WriteSources);
                package.XmlPart("customXml/itemProps1.xml", WriteSourceProperties);
                var item = new RelationshipList();
                item.Add(RelType + "customXmlProps", "itemProps1.xml");
                package.Relationships("customXml/_rels/item1.xml.rels", item);
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
            if (parts.HasNumbering)
            {
                package.Override("/word/numbering.xml", ContentType + "numbering+xml");
            }

            if (parts.Footnotes.Count > 0)
            {
                package.Override("/word/footnotes.xml", ContentType + "footnotes+xml");
            }

            if (_sources.Count > 0)
            {
                package.Override("/customXml/itemProps1.xml", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml");
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
    /// Numbers the headers and footers and gives every part, image, link, list, footnote and bookmark its id before
    /// anything is written.
    /// </summary>
    private Parts PlanParts()
    {
        var parts = new Parts();
        foreach (var section in _sections)
        {
            if (section.Watermark != null)
            {
                // The watermark lives in the headers, so the section needs one.
                _ = section.Header;
            }

            foreach (var container in Containers(section))
            {
                foreach (var paragraph in Paragraphs(container))
                {
                    parts.HasLists |= paragraph.List != null;
                    foreach (var run in paragraph.Runs)
                    {
                        parts.HasFootnoteRuns |= run.Footnote != null && container is DocxSection;
                    }
                }
            }
        }

        parts.HasNumbering = NumberHeadings || parts.HasLists;
        parts.Relationships.Add(RelationshipTypes.Styles, "styles.xml");
        parts.Relationships.Add(RelType + "settings", "settings.xml");
        parts.Relationships.Add(RelationshipTypes.Theme, "theme/theme1.xml");
        if (parts.HasNumbering)
        {
            parts.Relationships.Add(RelType + "numbering", "numbering.xml");
        }

        if (parts.HasFootnoteRuns)
        {
            parts.Relationships.Add(RelType + "footnotes", "footnotes.xml");
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
                if (!part.IsFooter && section.Watermark != null)
                {
                    parts.Watermarks[part] = section;
                }

                var own = new RelationshipList();
                parts.PartRelationships[part] = own;
                PlanRuns(part, own, parts, false);
            }

            parts.EvenPages |= section.EvenHeaderPart != null || section.EvenFooterPart != null;
        }

        foreach (var section in _sections)
        {
            PlanRuns(section, parts.Relationships, parts, true);
            foreach (var block in section.Blocks)
            {
                parts.HasContents |= block is DocxTableOfContents;
                parts.HasBibliography |= block is DocxBibliography;
            }
        }

        foreach (var source in parts.Citations)
        {
            if (!_sources.Contains(source))
            {
                throw new InvalidOperationException(Divit.Properties.Messages.Get(Divit.Properties.Strings.SourceNotInDocument));
            }
        }

        if (_sources.Count > 0)
        {
            parts.Relationships.Add(RelType + "customXml", "../customXml/item1.xml");
        }

        // Bookmarks for the paragraphs that links and references point at, in the order of the body.
        foreach (var section in _sections)
        {
            foreach (var paragraph in Paragraphs(section))
            {
                if (parts.Targets.Contains(paragraph))
                {
                    parts.Bookmarks[paragraph] = parts.Bookmarks.Count + parts.LabelBookmarks.Count;
                }

                if (parts.LabelTargets.Contains(paragraph))
                {
                    parts.LabelBookmarks[paragraph] = parts.Bookmarks.Count + parts.LabelBookmarks.Count;
                }
            }
        }

        if (parts.Bookmarks.Count != parts.Targets.Count || parts.LabelBookmarks.Count != parts.LabelTargets.Count)
        {
            throw new InvalidOperationException(Divit.Properties.Messages.Get(Divit.Properties.Strings.ReferenceTargetMissing));
        }

        parts.HeadingList = Headings(parts.HeadingNumbers);
        parts.Annotations = parts.Bookmarks.Count + parts.LabelBookmarks.Count;
        return parts;
    }

    /// <summary>
    /// A section and its headers and footers.
    /// </summary>
    private static IEnumerable<DocxContainer> Containers(DocxSection section)
    {
        yield return section;
        foreach (var part in new[] { section.DefaultHeaderPart, section.FirstHeaderPart, section.EvenHeaderPart, section.DefaultFooterPart, section.FirstFooterPart, section.EvenFooterPart })
        {
            if (part != null)
            {
                yield return part;
            }
        }
    }

    /// <summary>
    /// The paragraphs of a container in order, those in table cells (nested too) included.
    /// </summary>
    private static IEnumerable<DocxParagraph> Paragraphs(DocxContainer container)
    {
        foreach (var block in container.Blocks)
        {
            if (block is DocxParagraph paragraph)
            {
                yield return paragraph;
            }
            else if (block is DocxTable table)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var inner in Paragraphs(cell))
                        {
                            yield return inner;
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Plans what the runs of a part need: a media file and relationship per image (stored as it is), a
    /// relationship per web link, a numbering per list, the bookmarks links and references point at, and the
    /// footnotes of the body.
    /// </summary>
    private static void PlanRuns(DocxContainer container, RelationshipList relationships, Parts parts, bool body)
    {
        foreach (var paragraph in Paragraphs(container))
        {
            if (paragraph.List is DocxList list && !parts.ListIds.ContainsKey(list))
            {
                parts.ListIds[list] = parts.ListIds.Count + 2;
                parts.ListKinds.Add(list.Kind);
            }

            foreach (var run in paragraph.Runs)
            {
                if (run.Image != null)
                {
                    var name = "image" + (parts.Media.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "." + run.ImageInfo.Extension;
                    parts.Media.Add((name, run.Image));
                    parts.MediaTypes[run.ImageInfo.Extension] = run.ImageInfo.ContentType;
                    parts.ImageIds[run] = relationships.Add(RelationshipTypes.Image, "media/" + name);
                }
                else if (run.LinkUrl != null)
                {
                    parts.LinkIds[run] = relationships.Add(RelationshipTypes.Hyperlink, run.LinkUrl, true);
                    parts.HasLinks = true;
                }
                else if (run.LinkTarget != null)
                {
                    parts.Targets.Add(run.LinkTarget);
                    parts.HasLinks = true;
                }
                else if (run.ReferenceTarget != null)
                {
                    (run.ReferenceKind == DocxReferenceKind.Label ? parts.LabelTargets : parts.Targets).Add(run.ReferenceTarget);
                    parts.HasReferences = true;
                }
                else if (run.CheckBox != null)
                {
                    parts.CheckBoxParts.Add(container);
                }
                else if (run.Footnote != null && body)
                {
                    parts.FootnoteIds[run] = parts.Footnotes.Count + 1;
                    parts.Footnotes.Add(run);
                }
                else if (run.Math != null)
                {
                    parts.MathParts.Add(container);
                }
                else if (run.Citation != null)
                {
                    parts.Citations.Add(run.Citation);
                }
            }
        }
    }

    /// <summary>
    /// The parts of one save: relationships of the main part, the headers and footers with their ids, and the ids
    /// of images, links, lists, footnotes and bookmarks.
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
        /// The relationships of each header and footer part (its images and links).
        /// </summary>
        public Dictionary<DocxHeaderFooter, RelationshipList> PartRelationships { get; } = new Dictionary<DocxHeaderFooter, RelationshipList>();

        /// <summary>
        /// The headers that draw their section's watermark.
        /// </summary>
        public Dictionary<DocxHeaderFooter, DocxSection> Watermarks { get; } = new Dictionary<DocxHeaderFooter, DocxSection>();

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
        /// The relationship id of each web link, in the part that holds it.
        /// </summary>
        public Dictionary<DocxRun, string> LinkIds { get; } = new Dictionary<DocxRun, string>();

        /// <summary>
        /// The numbering id of each list; 1 is the heading outline.
        /// </summary>
        public Dictionary<DocxList, int> ListIds { get; } = new Dictionary<DocxList, int>();

        /// <summary>
        /// The kinds of list used, each with its own numbering definition.
        /// </summary>
        public SortedSet<DocxListKind> ListKinds { get; } = new SortedSet<DocxListKind>();

        /// <summary>
        /// Paragraphs that links and references point at as a whole, and captions whose label they show.
        /// </summary>
        public HashSet<DocxParagraph> Targets { get; } = new HashSet<DocxParagraph>();

        public HashSet<DocxParagraph> LabelTargets { get; } = new HashSet<DocxParagraph>();

        /// <summary>
        /// Bookmark ids around whole paragraphs and around the label of captions.
        /// </summary>
        public Dictionary<DocxParagraph, int> Bookmarks { get; } = new Dictionary<DocxParagraph, int>();

        public Dictionary<DocxParagraph, int> LabelBookmarks { get; } = new Dictionary<DocxParagraph, int>();

        /// <summary>
        /// The footnotes of the body in order, and the id of each.
        /// </summary>
        public List<DocxRun> Footnotes { get; } = new List<DocxRun>();

        public Dictionary<DocxRun, int> FootnoteIds { get; } = new Dictionary<DocxRun, int>();

        /// <summary>
        /// The parts (a section stands for the main part) that hold check boxes, which need the Word 2010 namespace.
        /// </summary>
        public HashSet<DocxContainer> CheckBoxParts { get; } = new HashSet<DocxContainer>();

        /// <summary>
        /// The headings in order with their numbers, and the number of each heading paragraph.
        /// </summary>
        public IReadOnlyList<(int Level, string Number, string Text)> HeadingList { get; set; } = Array.Empty<(int, string, string)>();

        public Dictionary<DocxParagraph, string> HeadingNumbers { get; } = new Dictionary<DocxParagraph, string>();

        /// <summary>
        /// Counter of drawing ids, which Word wants unique in the document.
        /// </summary>
        public int Drawings { get; set; }

        /// <summary>
        /// Counters of content control and shape ids, unique in the document.
        /// </summary>
        public int Controls { get; set; }

        public int Shapes { get; set; }

        /// <summary>
        /// Whether any section has its own even-page header or footer, which Word turns on for the whole document.
        /// </summary>
        public bool EvenPages { get; set; }

        /// <summary>
        /// Whether the document has a table of contents, which asks Word to update the fields on opening.
        /// </summary>
        public bool HasContents { get; set; }

        public bool HasLists { get; set; }

        /// <summary>
        /// Whether a numbering part is written: numbered headings or lists.
        /// </summary>
        public bool HasNumbering { get; set; }

        public bool HasLinks { get; set; }

        public bool HasReferences { get; set; }

        public bool HasFootnoteRuns { get; set; }

        public bool HasBibliography { get; set; }

        /// <summary>
        /// The parts that hold equations, which declare the math namespace.
        /// </summary>
        public HashSet<DocxContainer> MathParts { get; } = new HashSet<DocxContainer>();

        /// <summary>
        /// The sources cited anywhere.
        /// </summary>
        public HashSet<DocxSource> Citations { get; } = new HashSet<DocxSource>();

        /// <summary>
        /// Counter of revision ids, which continue after the bookmark ids.
        /// </summary>
        public int Annotations { get; set; }

        /// <summary>
        /// Whether Word should update the fields on opening: contents, references, citations, a bibliography.
        /// </summary>
        public bool UpdateFields => HasContents || HasReferences || Citations.Count > 0 || HasBibliography;
    }
}
