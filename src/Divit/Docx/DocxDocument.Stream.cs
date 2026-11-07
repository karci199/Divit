#nullable enable
namespace Divit.Docx;

using System;
using System.Xml;

/// <content>
/// Streamed sections: content written straight into the main part while it is saved.
/// </content>
public sealed partial class DocxDocument
{
    /// <summary>
    /// Starts a new section on a new page whose content is written while the document is saved, by
    /// <paramref name="write"/>, and never kept in memory: for very long reports, logs and tables of many thousand
    /// rows. Page setup, headers, footers and watermark are set on the returned section as usual.
    /// </summary>
    /// <remarks>
    /// The callback runs on every save; it must write the same content again for the same file. What a streamed
    /// section cannot hold is listed under <see cref="DocxStreamWriter"/>.
    /// </remarks>
    /// <param name="write">Writes the section's content, top to bottom.</param>
    /// <returns>The section, to set everything but its content.</returns>
    public DocxSection AddStreamedSection(Action<DocxStreamWriter> write)
    {
        var section = AddSection();
        section.Source = write ?? throw new ArgumentNullException(nameof(write));
        return section;
    }

    /// <summary>
    /// Whether any section is streamed: its content is unknown until it is written, so the main part declares the
    /// namespaces of equations and check boxes in advance.
    /// </summary>
    private bool HasStreamedSections()
    {
        foreach (var section in _sections)
        {
            if (section.Source != null)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Plans each block of a streamed section the moment it is complete, as the save plans stored ones, and writes
    /// it into the main part.
    /// </summary>
    private sealed class StreamSink : DocxBlockSink
    {
        private readonly DocxDocument _document;
        private readonly XmlWriter _w;
        private readonly DocxSection _section;
        private readonly Scope _scope;
        private TableState? _table;

        public StreamSink(DocxDocument document, XmlWriter w, DocxSection section, Scope scope)
        {
            _document = document;
            _w = w;
            _section = section;
            _scope = scope;
        }

        public override void Paragraph(DocxParagraph paragraph)
        {
            Plan(paragraph);
            _document.WriteParagraph(_w, paragraph, _scope);
        }

        public override void StartTable(DocxTable table) => _table = DocxDocument.StartTable(_w, table, _scope);

        public override void Row(DocxTableRow row)
        {
            foreach (var cell in row.Cells)
            {
                foreach (var paragraph in Paragraphs(cell))
                {
                    Plan(paragraph);
                }
            }

            _document.WriteRow(_w, row, _table!, _scope);
        }

        public override void EndTable()
        {
            _w.WriteEndElement();
            _table = null;
        }

        private void Plan(DocxParagraph paragraph)
        {
            var parts = _scope.Parts;
            foreach (var run in paragraph.Runs)
            {
                if (run.LinkTarget != null || run.ReferenceTarget != null)
                {
                    throw new InvalidOperationException(Divit.Properties.Messages.Get(Divit.Properties.Strings.ReferenceTargetMissing));
                }

                if (run.Citation != null && !_document._sources.Contains(run.Citation))
                {
                    throw new InvalidOperationException(Divit.Properties.Messages.Get(Divit.Properties.Strings.SourceNotInDocument));
                }
            }

            PlanParagraph(paragraph, _section, parts.Relationships, parts, true);
        }
    }

    /// <summary>
    /// Runs a streamed section's callback and writes what it adds.
    /// </summary>
    private void WriteStreamed(XmlWriter w, DocxSection section, Scope scope)
    {
        var writer = new DocxStreamWriter(new StreamSink(this, w, section, scope));
        section.Source!(writer);
        writer.Flush();
    }
}
