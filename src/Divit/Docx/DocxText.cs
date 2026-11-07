#nullable enable
namespace Divit.Docx;

using System;
using System.Collections.Generic;

/// <summary>
/// The look of a piece of text: font, size, emphasis, colour. An empty field takes the paragraph style's value.
/// Colours are <c>#RRGGBB</c>.
/// </summary>
public struct DocxFont
{
    /// <summary>
    /// Font name; the style's font when empty.
    /// </summary>
    public string? Name;

    /// <summary>
    /// Size in points; 0 keeps the style's size.
    /// </summary>
    public double Size;

    /// <summary>
    /// Bold text.
    /// </summary>
    public bool Bold;

    /// <summary>
    /// Italic text.
    /// </summary>
    public bool Italic;

    /// <summary>
    /// Underlined text.
    /// </summary>
    public bool Underline;

    /// <summary>
    /// Struck-through text.
    /// </summary>
    public bool Strikethrough;

    /// <summary>
    /// Capital letters, whatever was typed.
    /// </summary>
    public bool AllCaps;

    /// <summary>
    /// Small capital letters.
    /// </summary>
    public bool SmallCaps;

    /// <summary>
    /// Text colour; the style's colour when empty.
    /// </summary>
    public string? Color;

    /// <summary>
    /// Background behind the text (shading); none when empty.
    /// </summary>
    public string? Background;

    /// <summary>
    /// Superscript or subscript.
    /// </summary>
    public DocxScript Script;
}

/// <summary>
/// Raised or lowered text.
/// </summary>
public enum DocxScript
{
    /// <summary>
    /// On the line.
    /// </summary>
    None,

    /// <summary>
    /// Raised and small (m²).
    /// </summary>
    Superscript,

    /// <summary>
    /// Lowered and small (H₂O).
    /// </summary>
    Subscript,
}

/// <summary>
/// Horizontal alignment of a paragraph.
/// </summary>
public enum DocxAlignment
{
    /// <summary>
    /// Left (the style's alignment).
    /// </summary>
    Left,

    /// <summary>
    /// Centred.
    /// </summary>
    Center,

    /// <summary>
    /// Right.
    /// </summary>
    Right,

    /// <summary>
    /// Both edges straight.
    /// </summary>
    Justify,
}

/// <summary>
/// A value Word fills in when it lays out the document.
/// </summary>
public enum DocxField
{
    /// <summary>
    /// The page number.
    /// </summary>
    Page,

    /// <summary>
    /// The number of pages of the document.
    /// </summary>
    NumPages,

    /// <summary>
    /// The number of pages of the section.
    /// </summary>
    SectionPages,

    /// <summary>
    /// Today's date when Word prints or updates the field (<c>dd.MM.yyyy</c>).
    /// </summary>
    Date,

    /// <summary>
    /// The document title from its properties.
    /// </summary>
    Title,
}

/// <summary>
/// Where text lines up at a tab stop.
/// </summary>
public enum DocxTabAlignment
{
    /// <summary>
    /// Text starts at the stop.
    /// </summary>
    Left,

    /// <summary>
    /// Text is centred on the stop.
    /// </summary>
    Center,

    /// <summary>
    /// Text ends at the stop (a page number at the right margin).
    /// </summary>
    Right,
}

/// <summary>
/// The characters that fill the space before a tab stop.
/// </summary>
public enum DocxTabLeader
{
    /// <summary>
    /// Blank.
    /// </summary>
    None,

    /// <summary>
    /// Dots, as in a table of contents.
    /// </summary>
    Dot,

    /// <summary>
    /// A dashed line.
    /// </summary>
    Hyphen,

    /// <summary>
    /// A solid line.
    /// </summary>
    Underscore,
}

/// <summary>
/// The kind of a border line.
/// </summary>
public enum DocxLineStyle
{
    /// <summary>
    /// A single line.
    /// </summary>
    Single,

    /// <summary>
    /// Two lines.
    /// </summary>
    Double,

    /// <summary>
    /// Dots.
    /// </summary>
    Dotted,

    /// <summary>
    /// Dashes.
    /// </summary>
    Dashed,
}

/// <summary>
/// A border line: kind, width and colour.
/// </summary>
public readonly struct DocxBorder : IEquatable<DocxBorder>
{
    /// <summary>
    /// Creates a border line.
    /// </summary>
    /// <param name="color">Colour as <c>#RRGGBB</c>.</param>
    /// <param name="width">Width in points (0.25–6).</param>
    /// <param name="style">Kind of line.</param>
    public DocxBorder(string color, double width = 0.5, DocxLineStyle style = DocxLineStyle.Single)
    {
        Color = color;
        Width = width;
        Style = style;
    }

    /// <summary>
    /// Colour as <c>#RRGGBB</c>.
    /// </summary>
    public string Color { get; }

    /// <summary>
    /// Width in points.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// Kind of line.
    /// </summary>
    public DocxLineStyle Style { get; }

    /// <inheritdoc />
    public bool Equals(DocxBorder other) => Color == other.Color && Width.Equals(other.Width) && Style == other.Style;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is DocxBorder other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => ((Color?.GetHashCode() ?? 0) * 397) ^ Width.GetHashCode() ^ ((int)Style << 24);
}

/// <summary>
/// Names of the paragraph styles every Divit document defines. Word shows the built-in ones under their names in
/// the user's language (Heading 1 is "Başlık 1" in a Turkish Word).
/// </summary>
public static class DocxStyle
{
    /// <summary>
    /// Body text.
    /// </summary>
    public const string Normal = "Normal";

    /// <summary>
    /// The large title of a cover page.
    /// </summary>
    public const string Title = "Title";

    /// <summary>
    /// The line under the title.
    /// </summary>
    public const string Subtitle = "Subtitle";

    /// <summary>
    /// The caption of a figure or table.
    /// </summary>
    public const string Caption = "Caption";

    /// <summary>
    /// Body text without space after it: addresses, headers and footers.
    /// </summary>
    public const string NoSpacing = "NoSpacing";

    /// <summary>
    /// An item of a bulleted or numbered list (<see cref="DocxList.AddItem"/>).
    /// </summary>
    public const string ListParagraph = "ListParagraph";

    /// <summary>
    /// The heading of a level, 1–6.
    /// </summary>
    /// <param name="level">The level.</param>
    /// <returns>The style name.</returns>
    public static string Heading(int level) => "Heading" + Math.Max(1, Math.Min(6, level)).ToString(System.Globalization.CultureInfo.InvariantCulture);
}

/// <summary>
/// One piece of a paragraph: text, a tab, a line break or a field, with its font.
/// </summary>
internal sealed class DocxRun
{
    public string? Text { get; set; }

    public DocxField? Field { get; set; }

    public bool Tab { get; set; }

    public bool LineBreak { get; set; }

    public bool PageBreak { get; set; }

    public DocxFont? Font { get; set; }

    /// <summary>
    /// A field written by its code, such as the number of a caption (<c>SEQ Figure</c>), with the text shown until
    /// Word updates it.
    /// </summary>
    public string? Code { get; set; }

    public string? Result { get; set; }

    /// <summary>
    /// An image in the line: its bytes, format, size in EMUs and alternative text.
    /// </summary>
    public byte[]? Image { get; set; }

    public Divit.Packaging.ImageInfo ImageInfo { get; set; }

    public long ImageWidth { get; set; }

    public long ImageHeight { get; set; }

    public string? Description { get; set; }

    /// <summary>
    /// A link around the text: to a web address, or to a paragraph of the document.
    /// </summary>
    public string? LinkUrl { get; set; }

    public DocxParagraph? LinkTarget { get; set; }

    /// <summary>
    /// A cross-reference field: the paragraph it points at and what it shows of it.
    /// </summary>
    public DocxParagraph? ReferenceTarget { get; set; }

    public DocxReferenceKind ReferenceKind { get; set; }

    /// <summary>
    /// A check box that readers can tick; whether it starts ticked.
    /// </summary>
    public bool? CheckBox { get; set; }

    /// <summary>
    /// A footnote: the mark in the text, and the note's text at the foot of the page.
    /// </summary>
    public string? Footnote { get; set; }
}

/// <summary>
/// The kind of a caption: what it numbers.
/// </summary>
public enum DocxCaptionKind
{
    /// <summary>
    /// A figure: "Figure 1".
    /// </summary>
    Figure,

    /// <summary>
    /// A table: "Table 1".
    /// </summary>
    Table,
}

/// <summary>
/// A tab stop of a paragraph.
/// </summary>
internal readonly struct DocxTabStop
{
    public DocxTabStop(double position, DocxTabAlignment alignment, DocxTabLeader leader)
    {
        Position = position;
        Alignment = alignment;
        Leader = leader;
    }

    /// <summary>
    /// Position in centimetres from the left margin.
    /// </summary>
    public double Position { get; }

    public DocxTabAlignment Alignment { get; }

    public DocxTabLeader Leader { get; }
}

/// <summary>
/// Something in the flow of a document: a paragraph, a table, a table of contents.
/// </summary>
public abstract class DocxBlock
{
    internal DocxBlock()
    {
    }
}

/// <summary>
/// A paragraph: pieces of text with their fonts, and the paragraph's own layout (alignment, spacing, indents,
/// shading, border). Methods return the paragraph, so pieces chain: <c>p.Add("Total: ").Add("12", bold)</c>.
/// </summary>
public sealed class DocxParagraph : DocxBlock
{
    private readonly List<DocxRun> _runs = new List<DocxRun>();
    private readonly List<DocxTabStop> _tabs = new List<DocxTabStop>();

    internal DocxParagraph(string style)
    {
        Style = style;
    }

    /// <summary>
    /// The paragraph style, for example <see cref="DocxStyle.Normal"/> or <c>DocxStyle.Heading(2)</c>.
    /// </summary>
    public string Style { get; set; }

    /// <summary>
    /// Horizontal alignment; the style's when empty.
    /// </summary>
    public DocxAlignment? Alignment { get; set; }

    /// <summary>
    /// Space above the paragraph in points; the style's when empty.
    /// </summary>
    public double? SpaceBefore { get; set; }

    /// <summary>
    /// Space below the paragraph in points; the style's when empty.
    /// </summary>
    public double? SpaceAfter { get; set; }

    /// <summary>
    /// Line spacing as a multiple of single spacing (1, 1.15, 1.5, 2); the style's when empty.
    /// </summary>
    public double? LineSpacing { get; set; }

    /// <summary>
    /// Indent from the left margin in centimetres.
    /// </summary>
    public double LeftIndent { get; set; }

    /// <summary>
    /// Indent from the right margin in centimetres.
    /// </summary>
    public double RightIndent { get; set; }

    /// <summary>
    /// Extra indent of the first line in centimetres; negative for a hanging indent.
    /// </summary>
    public double FirstLineIndent { get; set; }

    /// <summary>
    /// Keeps the paragraph on the same page as the next one (headings, captions above tables).
    /// </summary>
    public bool KeepWithNext { get; set; }

    /// <summary>
    /// Keeps the lines of the paragraph on one page.
    /// </summary>
    public bool KeepTogether { get; set; }

    /// <summary>
    /// Starts the paragraph on a new page.
    /// </summary>
    public bool PageBreakBefore { get; set; }

    /// <summary>
    /// Background colour of the paragraph, edge to edge (a note box); none when empty.
    /// </summary>
    public string? Shading { get; set; }

    /// <summary>
    /// A frame around the paragraph; none when empty.
    /// </summary>
    public DocxBorder? Border { get; set; }

    internal IReadOnlyList<DocxRun> Runs => _runs;

    internal IReadOnlyList<DocxTabStop> TabStops => _tabs;

    /// <summary>
    /// The plain text of the paragraph, fields left out.
    /// </summary>
    public string Text
    {
        get
        {
            var text = new System.Text.StringBuilder();
            foreach (var run in _runs)
            {
                text.Append(run.Text);
                if (run.Tab) { text.Append('\t'); }
                if (run.LineBreak) { text.Append('\n'); }
            }

            return text.ToString();
        }
    }

    /// <summary>
    /// Adds text; a line feed in it breaks the line, a tab moves to the next tab stop.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <param name="font">Its font; the style's when empty.</param>
    /// <returns>This paragraph.</returns>
    public DocxParagraph Add(string text, DocxFont? font = null)
    {
        var start = 0;
        for (var i = 0; i <= text.Length; i++)
        {
            if (i == text.Length || text[i] == '\n' || text[i] == '\t')
            {
                if (i > start)
                {
                    _runs.Add(new DocxRun { Text = text.Substring(start, i - start).Replace("\r", string.Empty), Font = font });
                }

                if (i < text.Length)
                {
                    _runs.Add(new DocxRun { LineBreak = text[i] == '\n', Tab = text[i] == '\t', Font = font });
                }

                start = i + 1;
            }
        }

        return this;
    }

    /// <summary>
    /// Adds a field Word fills in: the page number, the page count, the date.
    /// </summary>
    /// <param name="field">The field.</param>
    /// <param name="font">Its font; the style's when empty.</param>
    /// <returns>This paragraph.</returns>
    public DocxParagraph AddField(DocxField field, DocxFont? font = null)
    {
        _runs.Add(new DocxRun { Field = field, Font = font });
        return this;
    }

    /// <summary>
    /// Moves to the next tab stop.
    /// </summary>
    /// <returns>This paragraph.</returns>
    public DocxParagraph AddTab()
    {
        _runs.Add(new DocxRun { Tab = true });
        return this;
    }

    /// <summary>
    /// Sets a tab stop: where text after a tab lines up.
    /// </summary>
    /// <param name="position">Position in centimetres from the left margin.</param>
    /// <param name="alignment">How the text lines up there.</param>
    /// <param name="leader">What fills the space before it.</param>
    /// <returns>This paragraph.</returns>
    public DocxParagraph AddTabStop(double position, DocxTabAlignment alignment = DocxTabAlignment.Left, DocxTabLeader leader = DocxTabLeader.None)
    {
        _tabs.Add(new DocxTabStop(position, alignment, leader));
        return this;
    }

    /// <summary>
    /// Adds an image (PNG, JPEG or GIF) in the line, for example a figure in a paragraph of its own or a logo in a
    /// header. With neither size given it takes the image's own size at 96 dpi; with one, it keeps the proportions.
    /// </summary>
    /// <param name="image">Bytes of the image file.</param>
    /// <param name="width">Width in centimetres; 0 to follow the height or the image.</param>
    /// <param name="height">Height in centimetres; 0 to follow the width or the image.</param>
    /// <param name="description">Alternative text for screen readers.</param>
    /// <returns>This paragraph.</returns>
    /// <exception cref="ArgumentException">The bytes are no PNG, JPEG or GIF, or the size is not given and cannot be
    /// read from them.</exception>
    public DocxParagraph AddImage(byte[] image, double width = 0, double height = 0, string? description = null)
    {
        var info = Divit.Packaging.ImageInfo.Read(image)
            ?? throw new ArgumentException(Divit.Properties.Messages.Get(Divit.Properties.Strings.ImageFormatUnknown), nameof(image));
        const double EmuPerCm = 360000;
        const double EmuPerPixel = 9525;
        double cx, cy;
        if (width > 0 && height > 0)
        {
            cx = width * EmuPerCm;
            cy = height * EmuPerCm;
        }
        else if (info.Width <= 0 || info.Height <= 0)
        {
            throw new ArgumentException(Divit.Properties.Messages.Get(Divit.Properties.Strings.ImageSizeUnknown), nameof(image));
        }
        else if (width > 0)
        {
            cx = width * EmuPerCm;
            cy = cx * info.Height / info.Width;
        }
        else if (height > 0)
        {
            cy = height * EmuPerCm;
            cx = cy * info.Width / info.Height;
        }
        else
        {
            cx = info.Width * EmuPerPixel;
            cy = info.Height * EmuPerPixel;
        }

        _runs.Add(new DocxRun { Image = image, ImageInfo = info, ImageWidth = (long)Math.Round(cx), ImageHeight = (long)Math.Round(cy), Description = description });
        return this;
    }

    /// <summary>
    /// Adds a link to a web page, a file or an e-mail address (<c>mailto:</c>); Word shows it in the link style.
    /// </summary>
    /// <param name="text">The text shown.</param>
    /// <param name="url">The address, for example <c>https://example.com</c>.</param>
    /// <param name="font">Its font; the link style's when empty.</param>
    /// <returns>This paragraph.</returns>
    /// <exception cref="ArgumentException">The address is empty.</exception>
    public DocxParagraph AddLink(string text, string url, DocxFont? font = null)
    {
        if (string.IsNullOrEmpty(url))
        {
            throw new ArgumentException(Divit.Properties.Messages.Get(Divit.Properties.Strings.LinkEmpty), nameof(url));
        }

        _runs.Add(new DocxRun { Text = text, LinkUrl = url, Font = font });
        return this;
    }

    /// <summary>
    /// Adds a link that jumps to a paragraph of the document, such as a heading or a caption.
    /// </summary>
    /// <param name="text">The text shown.</param>
    /// <param name="target">The paragraph to jump to; it must be in the body when the document is saved.</param>
    /// <param name="font">Its font; the link style's when empty.</param>
    /// <returns>This paragraph.</returns>
    public DocxParagraph AddLink(string text, DocxParagraph target, DocxFont? font = null)
    {
        _runs.Add(new DocxRun { Text = text, LinkTarget = target ?? throw new ArgumentNullException(nameof(target)), Font = font });
        return this;
    }

    /// <summary>
    /// Adds a cross-reference to a paragraph: its text, its heading number, its page, or a caption's label and number
    /// ("see Figure 3 on page 7"). Word updates it with the fields, so it follows when headings and captions move;
    /// until then it shows the values Divit knows (the page as 1). A click on it jumps to the paragraph.
    /// </summary>
    /// <param name="target">The paragraph it points at; it must be in the body when the document is saved.</param>
    /// <param name="kind">What it shows.</param>
    /// <param name="font">Its font; the paragraph's when empty.</param>
    /// <returns>This paragraph.</returns>
    /// <exception cref="ArgumentException"><see cref="DocxReferenceKind.Label"/> for a paragraph that is no
    /// caption.</exception>
    public DocxParagraph AddReference(DocxParagraph target, DocxReferenceKind kind = DocxReferenceKind.Text, DocxFont? font = null)
    {
        if (target == null)
        {
            throw new ArgumentNullException(nameof(target));
        }

        if (kind == DocxReferenceKind.Label && target.Caption == null)
        {
            throw new ArgumentException(Divit.Properties.Messages.Get(Divit.Properties.Strings.ReferenceNotCaption), nameof(target));
        }

        _runs.Add(new DocxRun { ReferenceTarget = target, ReferenceKind = kind, Font = font });
        return this;
    }

    /// <summary>
    /// Adds a check box readers can tick with a click (a Word 2010 content control); older readers show its box.
    /// </summary>
    /// <param name="isChecked">Whether it starts ticked.</param>
    /// <param name="font">Its size and colour; the paragraph's when empty.</param>
    /// <returns>This paragraph.</returns>
    public DocxParagraph AddCheckBox(bool isChecked = false, DocxFont? font = null)
    {
        _runs.Add(new DocxRun { CheckBox = isChecked, Font = font });
        return this;
    }

    /// <summary>
    /// Adds a footnote: a number here, and the note at the foot of the page. Footnotes belong to the body; in a
    /// header or footer they are left out.
    /// </summary>
    /// <param name="text">The note's text.</param>
    /// <returns>This paragraph.</returns>
    public DocxParagraph AddFootnote(string text)
    {
        _runs.Add(new DocxRun { Footnote = text ?? string.Empty });
        return this;
    }

    /// <summary>
    /// The list the paragraph is an item of, and its depth.
    /// </summary>
    internal DocxList? List { get; set; }

    internal int ListLevel { get; set; }

    /// <summary>
    /// A caption's label and number ("Figure 3"), which a <see cref="DocxReferenceKind.Label"/> reference shows.
    /// </summary>
    internal string? CaptionLabel { get; set; }

    internal void AddPageBreakRun() => _runs.Add(new DocxRun { PageBreak = true });

    internal void AddFieldRun(string code, string result) => _runs.Add(new DocxRun { Code = code, Result = result });

    /// <summary>
    /// The kind of caption the paragraph is, when it is one (<see cref="DocxDocument.AddCaption"/>).
    /// </summary>
    internal DocxCaptionKind? Caption { get; set; }

    /// <summary>
    /// The caption's text without its label and number, as the list of figures shows it.
    /// </summary>
    internal string? CaptionText { get; set; }
}

/// <summary>
/// A table of contents (<see cref="DocxDocument.AddTableOfContents"/>): Word fills in the page numbers when the file
/// opens; until then it lists the headings without them.
/// </summary>
public sealed class DocxTableOfContents : DocxBlock
{
    internal DocxTableOfContents(int levels)
    {
        Levels = levels;
    }

    /// <summary>
    /// Deepest heading level listed (1–6).
    /// </summary>
    public int Levels { get; }

    /// <summary>
    /// The caption kind a list of figures or tables lists; <c>null</c> for the headings.
    /// </summary>
    internal DocxCaptionKind? Captions { get; set; }
}

/// <summary>
/// Something that holds paragraphs and tables: a section of the document, a header, a footer.
/// </summary>
public abstract class DocxContainer
{
    private readonly List<DocxBlock> _blocks = new List<DocxBlock>();

    internal DocxContainer()
    {
    }

    /// <summary>
    /// The paragraphs and tables in order.
    /// </summary>
    public IReadOnlyList<DocxBlock> Blocks => _blocks;

    /// <summary>
    /// Adds a paragraph.
    /// </summary>
    /// <param name="text">Its text; none when empty.</param>
    /// <param name="font">Font of the text; the style's when empty.</param>
    /// <param name="style">Paragraph style; <see cref="DocxStyle.Normal"/> when empty.</param>
    /// <returns>The paragraph, to add more text or set its layout.</returns>
    public DocxParagraph AddParagraph(string? text = null, DocxFont? font = null, string? style = null)
    {
        var paragraph = new DocxParagraph(style ?? DefaultStyle);
        if (!string.IsNullOrEmpty(text))
        {
            paragraph.Add(text!, font);
        }

        _blocks.Add(paragraph);
        return paragraph;
    }

    /// <summary>
    /// The style of a paragraph added without one.
    /// </summary>
    internal virtual string DefaultStyle => DocxStyle.Normal;

    /// <summary>
    /// Adds a table.
    /// </summary>
    /// <param name="columns">Number of columns, 1–63.</param>
    /// <returns>The table, to add rows and set widths.</returns>
    public DocxTable AddTable(int columns)
    {
        var table = new DocxTable(columns);
        _blocks.Add(table);
        return table;
    }

    /// <summary>
    /// Starts a bulleted or numbered list; its items follow where they are added.
    /// </summary>
    /// <param name="kind">Bullets or numbers.</param>
    /// <returns>The list, to add items.</returns>
    public DocxList AddList(DocxListKind kind = DocxListKind.Bullet) => new DocxList(this, kind);

    internal void AddBlock(DocxBlock block) => _blocks.Add(block);
}
