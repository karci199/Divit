#nullable enable
namespace Divit.Docx;

using System;

/// <summary>
/// Writes the content of a streamed section (<see cref="DocxDocument.AddStreamedSection"/>) while the document is
/// saved: each paragraph or table goes into the file as soon as the next one starts, and the rows of a table as soon
/// as the next row starts, so even very long documents take little memory.
/// </summary>
/// <remarks>
/// <para>Everything added here works as in a stored section: text and fonts, headings, lists, tables with merges,
/// images, links to web pages, check boxes, footnotes, equations, citations and tracked changes. A paragraph or
/// row can be changed until the next one is added; set a table's widths and look before its second row.</para>
/// <para>Links and cross-references to paragraphs need their bookmarks before the body is written, so they are not
/// available in a streamed section, neither pointing out of it nor into it; captions are numbered by the document,
/// so add them in stored sections. Word adds the headings of streamed sections to a table of contents when it
/// updates the fields on opening.</para>
/// </remarks>
public sealed class DocxStreamWriter : DocxContainer
{
    private readonly DocxBlockSink _sink;
    private DocxBlock? _pending;
    private bool _tableStarted;

    internal DocxStreamWriter(DocxBlockSink sink)
    {
        _sink = sink;
    }

    /// <summary>
    /// Adds a heading; with <see cref="DocxDocument.NumberHeadings"/> it gets its number.
    /// </summary>
    /// <param name="text">The heading text.</param>
    /// <param name="level">Level 1–6.</param>
    /// <returns>The paragraph.</returns>
    public DocxParagraph AddHeading(string text, int level = 1) => AddParagraph(text, null, DocxStyle.Heading(level));

    /// <summary>
    /// Continues on a new page.
    /// </summary>
    public void AddPageBreak() => AddParagraph().AddPageBreakRun();

    /// <summary>
    /// Adds an equation set apart and centred in a paragraph of its own.
    /// </summary>
    /// <param name="math">The equation.</param>
    /// <returns>The paragraph.</returns>
    public DocxParagraph AddEquation(DocxMath math)
    {
        var paragraph = AddParagraph();
        paragraph.AddDisplayMath(math ?? throw new ArgumentNullException(nameof(math)));
        return paragraph;
    }

    /// <summary>
    /// Writes the block before this one and keeps this one open for changes.
    /// </summary>
    internal override void Append(DocxBlock block)
    {
        Flush();
        _pending = block;
        if (block is DocxTable table)
        {
            _tableStarted = false;
            table.BeforeRow = WriteRows;
        }
    }

    /// <summary>
    /// Writes what is still open; called when the callback returns.
    /// </summary>
    internal void Flush()
    {
        switch (_pending)
        {
            case DocxParagraph paragraph:
                _sink.Paragraph(paragraph);
                break;
            case DocxTable table:
                WriteRows(table);
                table.BeforeRow = null;
                _sink.EndTable();
                break;
        }

        _pending = null;
    }

    /// <summary>
    /// Writes the rows of the open table so far, opening it first, and forgets them.
    /// </summary>
    private void WriteRows(DocxTable table)
    {
        if (!_tableStarted)
        {
            _sink.StartTable(table);
            _tableStarted = true;
        }

        foreach (var row in table.Rows)
        {
            _sink.Row(row);
        }

        table.ClearRows();
    }
}

/// <summary>
/// Where a streamed section's blocks are written.
/// </summary>
internal abstract class DocxBlockSink
{
    public abstract void Paragraph(DocxParagraph paragraph);

    public abstract void StartTable(DocxTable table);

    public abstract void Row(DocxTableRow row);

    public abstract void EndTable();
}
