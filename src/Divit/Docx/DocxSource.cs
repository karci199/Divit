#nullable enable
namespace Divit.Docx;

using System;
using System.Collections.Generic;

/// <summary>
/// The kind of a source, as Word's source manager names it.
/// </summary>
public enum DocxSourceKind
{
    /// <summary>
    /// A book.
    /// </summary>
    Book,

    /// <summary>
    /// A report, a standard, a specification or a data sheet.
    /// </summary>
    Report,

    /// <summary>
    /// An article in a journal.
    /// </summary>
    JournalArticle,

    /// <summary>
    /// A paper in conference proceedings.
    /// </summary>
    ConferenceProceedings,

    /// <summary>
    /// A web site.
    /// </summary>
    InternetSite,

    /// <summary>
    /// Anything else.
    /// </summary>
    Misc,
}

/// <summary>
/// How citations and the bibliography are written; Word formats them again in this style when it updates the
/// fields, and readers can switch it under References.
/// </summary>
public enum DocxBibliographyStyle
{
    /// <summary>
    /// IEEE: numbered, "[1]".
    /// </summary>
    Ieee,

    /// <summary>
    /// APA: author and year, "(Balanis, 2016)".
    /// </summary>
    Apa,
}

/// <summary>
/// A source of the bibliography (<see cref="DocxDocument.AddSource"/>): a book, a report, a standard, an article, a
/// web site. Word lists it in its source manager; <see cref="DocxParagraph.AddCitation"/> cites it and
/// <see cref="DocxDocument.AddBibliography"/> lists the cited sources.
/// </summary>
public sealed class DocxSource
{
    private readonly List<(string Last, string First)> _authors = new List<(string, string)>();

    internal DocxSource(string tag, DocxSourceKind kind, string title)
    {
        Tag = tag;
        Kind = kind;
        Title = title;
    }

    /// <summary>
    /// The short name Word cites the source by, unique in the document.
    /// </summary>
    public string Tag { get; }

    /// <summary>
    /// The kind of source.
    /// </summary>
    public DocxSourceKind Kind { get; }

    /// <summary>
    /// Title.
    /// </summary>
    public string Title { get; }

    /// <summary>
    /// Year of publication.
    /// </summary>
    public string? Year { get; set; }

    /// <summary>
    /// An organisation as the author, such as "IEEE" or "Department of Defense", instead of persons.
    /// </summary>
    public string? Corporate { get; set; }

    /// <summary>
    /// Publisher or issuing organisation.
    /// </summary>
    public string? Publisher { get; set; }

    /// <summary>
    /// City of the publisher.
    /// </summary>
    public string? City { get; set; }

    /// <summary>
    /// Name of the journal or of the conference.
    /// </summary>
    public string? Journal { get; set; }

    /// <summary>
    /// Volume of a journal.
    /// </summary>
    public string? Volume { get; set; }

    /// <summary>
    /// Issue of a journal.
    /// </summary>
    public string? Issue { get; set; }

    /// <summary>
    /// Pages, for example <c>1–12</c>.
    /// </summary>
    public string? Pages { get; set; }

    /// <summary>
    /// Edition of a book, for example <c>4th</c>.
    /// </summary>
    public string? Edition { get; set; }

    /// <summary>
    /// Web address.
    /// </summary>
    public string? Url { get; set; }

    /// <summary>
    /// The authors in order, family name and given names.
    /// </summary>
    public IReadOnlyList<(string Last, string First)> Authors => _authors;

    /// <summary>
    /// Adds an author.
    /// </summary>
    /// <param name="last">Family name.</param>
    /// <param name="first">Given names.</param>
    /// <returns>This source.</returns>
    public DocxSource AddAuthor(string last, string first = "")
    {
        _authors.Add((last ?? throw new ArgumentNullException(nameof(last)), first ?? string.Empty));
        return this;
    }

    /// <summary>
    /// Position in the document's source list, from 1: the IEEE number shown until Word updates the fields.
    /// </summary>
    internal int Number { get; set; }
}

/// <summary>
/// The bibliography (<see cref="DocxDocument.AddBibliography"/>): a field Word fills with the document's sources in
/// its bibliography style; until then it lists them as Divit writes them.
/// </summary>
public sealed class DocxBibliography : DocxBlock
{
    internal DocxBibliography()
    {
    }
}
