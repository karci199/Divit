#nullable enable
namespace Divit.Docx;

using System;

/// <summary>
/// The kind of a list: what stands before its items.
/// </summary>
public enum DocxListKind
{
    /// <summary>
    /// Bullets: •, then – and » on deeper levels.
    /// </summary>
    Bullet,

    /// <summary>
    /// Numbers: 1., then a) and i. on deeper levels.
    /// </summary>
    Numbered,
}

/// <summary>
/// A bulleted or numbered list (<see cref="DocxContainer.AddList"/>). Its items are paragraphs in the flow of the
/// container, so other paragraphs may stand between them; a numbered list counts on across them and every list
/// starts again at 1.
/// </summary>
public sealed class DocxList
{
    private readonly DocxContainer _container;

    internal DocxList(DocxContainer container, DocxListKind kind)
    {
        _container = container;
        Kind = kind;
    }

    /// <summary>
    /// Bullets or numbers.
    /// </summary>
    public DocxListKind Kind { get; }

    /// <summary>
    /// Adds an item at the end of the container.
    /// </summary>
    /// <param name="text">The item's text; none when empty.</param>
    /// <param name="level">Depth, 0 (outermost) to 8.</param>
    /// <param name="font">Font of the text; the style's when empty.</param>
    /// <returns>The item's paragraph, to add more text.</returns>
    public DocxParagraph AddItem(string? text = null, int level = 0, DocxFont? font = null)
    {
        var paragraph = _container.AddParagraph(text, font, DocxStyle.ListParagraph);
        paragraph.List = this;
        paragraph.ListLevel = Math.Max(0, Math.Min(8, level));
        return paragraph;
    }
}

/// <summary>
/// What a cross-reference (<see cref="DocxParagraph.AddReference"/>) shows of the paragraph it points at. Word
/// keeps it up to date when the fields are updated.
/// </summary>
public enum DocxReferenceKind
{
    /// <summary>
    /// The paragraph's text: a heading's title.
    /// </summary>
    Text,

    /// <summary>
    /// A heading's number: "4.1".
    /// </summary>
    Number,

    /// <summary>
    /// The page the paragraph is on.
    /// </summary>
    Page,

    /// <summary>
    /// A caption's label and number: "Figure 3".
    /// </summary>
    Label,
}

/// <summary>
/// What a protected document (<see cref="DocxDocument.Protect"/>) still lets readers change.
/// </summary>
public enum DocxProtection
{
    /// <summary>
    /// Nothing: the document is read only.
    /// </summary>
    ReadOnly,

    /// <summary>
    /// Only the form controls, such as check boxes.
    /// </summary>
    Forms,

    /// <summary>
    /// Only comments.
    /// </summary>
    Comments,

    /// <summary>
    /// Everything, but every change is tracked.
    /// </summary>
    TrackedChanges,
}
