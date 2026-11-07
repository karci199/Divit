#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using System;
using System.Xml;

/// <content>
/// Streamed sheets: rows written straight into the sheet part while it is saved.
/// </content>
public sealed partial class XlsxWorkbook
{
    /// <summary>
    /// Adds a sheet whose rows are written while the workbook is saved, by <paramref name="writeRows"/>, and never kept
    /// in memory: for millions of rows. Everything else of the sheet (column widths, frozen panes, merges, conditional
    /// formats, validation, links, comments, protection, print settings) is set on the returned sheet as usual.
    /// </summary>
    /// <remarks>
    /// <para>The cells of a streamed sheet come only from the callback: <see cref="XlsxSheet.Set(int, int, string)"/>,
    /// <see cref="XlsxSheet.Format(int, int, int, int, XlsxStyleEdit)"/>, row heights and tables are not available on
    /// it, and charts on it carry no cached values (Excel reads the cells when it opens the file).</para>
    /// <para>The callback runs on every save; it must write the same rows again for the same file.</para>
    /// <para>Repeated texts are stored once in the shared string table, which stays in memory until the save ends. For
    /// millions of different texts (ids, free text) set <paramref name="inlineStrings"/>: each text is then written
    /// into its cell, which keeps memory flat at the cost of a larger file.</para>
    /// </remarks>
    /// <param name="name">Tab name, by the rules of <see cref="AddSheet"/>.</param>
    /// <param name="writeRows">Writes the rows, top to bottom.</param>
    /// <param name="inlineStrings">Whether texts go into their cells instead of the shared string table.</param>
    /// <returns>The sheet, to set everything but its cells.</returns>
    public XlsxSheet AddStreamedSheet(string name, Action<XlsxRowWriter> writeRows, bool inlineStrings = false)
    {
        var sheet = AddSheet(name);
        sheet.RowSource = writeRows;
        sheet.InlineStrings = inlineStrings;
        return sheet;
    }

    /// <summary>
    /// Writes the cells of a row writer as worksheet XML, with the same style and string tables as stored sheets.
    /// </summary>
    private sealed class StreamSink : XlsxRowSink
    {
        private readonly XlsxWorkbook _book;
        private readonly XmlWriter _w;
        private readonly StyleTable _styles;
        private readonly StringTable _strings;
        private readonly char[] _buffer;
        private readonly bool _inline;

        public StreamSink(XlsxWorkbook book, XmlWriter w, StyleTable styles, StringTable strings, char[] buffer, bool inline)
        {
            _book = book;
            _w = w;
            _styles = styles;
            _strings = strings;
            _buffer = buffer;
            _inline = inline;
        }

        public override void StartRow(int row, double? height, bool hidden, int outlineLevel)
        {
            _w.WriteStartElement("row", MainNs);
            WriteNumberAttribute(_w, "r", row, _buffer);
            if (height is double points)
            {
                _w.WriteAttributeString("ht", Num(points));
                _w.WriteAttributeString("customHeight", "1");
            }

            if (hidden) { _w.WriteAttributeString("hidden", "1"); }
            if (outlineLevel > 0) { _w.WriteAttributeString("outlineLevel", Int(outlineLevel)); }
        }

        public override void EndRow() => _w.WriteEndElement();

        public override void Empty(int row, int column, in XlsxStyle style)
        {
            var index = _styles.Index(in style);
            if (index != 0)
            {
                Start(row, column, index);
                _w.WriteEndElement();
            }
        }

        public override void Text(int row, int column, string text, in XlsxStyle style)
        {
            Start(row, column, _styles.Index(in style));
            if (_inline)
            {
                _w.WriteAttributeString("t", "inlineStr");
                _w.WriteStartElement("is", MainNs);
                StringTable.Text(_w, text);
                _w.WriteEndElement();
            }
            else
            {
                _w.WriteAttributeString("t", "s");
                _w.WriteStartElement("v", MainNs);
                _w.WriteChars(_buffer, 0, OfficeText.FormatInt(_buffer, 0, _strings.Index(text)));
                _w.WriteEndElement();
            }

            _w.WriteEndElement();
        }

        public override void Number(int row, int column, double number, in XlsxStyle style)
        {
            Start(row, column, _styles.Index(in style));
            _w.WriteStartElement("v", MainNs);
            _w.WriteChars(_buffer, 0, OfficeText.FormatDouble(_buffer, number));
            _w.WriteEndElement();
            _w.WriteEndElement();
        }

        public override void Bool(int row, int column, bool value, in XlsxStyle style)
        {
            Start(row, column, _styles.Index(in style));
            _w.WriteAttributeString("t", "b");
            _w.WriteElementString("v", MainNs, value ? "1" : "0");
            _w.WriteEndElement();
        }

        public override void Rich(int row, int column, XlsxRichText text)
        {
            Start(row, column, 0);
            if (_inline)
            {
                _w.WriteAttributeString("t", "inlineStr");
                _w.WriteStartElement("is", MainNs);
                _strings.Runs(_w, text);
                _w.WriteEndElement();
            }
            else
            {
                _w.WriteAttributeString("t", "s");
                _w.WriteStartElement("v", MainNs);
                _w.WriteChars(_buffer, 0, OfficeText.FormatInt(_buffer, 0, _strings.Index(text)));
                _w.WriteEndElement();
            }

            _w.WriteEndElement();
        }

        public override void Formula(int row, int column, string formula, double? result, in XlsxStyle style)
        {
            _book._formulas = true;
            Start(row, column, _styles.Index(in style));
            _w.WriteElementString("f", MainNs, OfficeText.Clean(formula));
            if (result is double value && !double.IsNaN(value) && !double.IsInfinity(value))
            {
                _w.WriteStartElement("v", MainNs);
                _w.WriteChars(_buffer, 0, OfficeText.FormatDouble(_buffer, value));
                _w.WriteEndElement();
            }

            _w.WriteEndElement();
        }

        private void Start(int row, int column, int style)
        {
            _w.WriteStartElement("c", MainNs);
            _w.WriteStartAttribute("r");
            _w.WriteChars(_buffer, 0, FormatCellName(_buffer, row, column));
            _w.WriteEndAttribute();
            if (style != 0)
            {
                WriteNumberAttribute(_w, "s", style, _buffer);
            }
        }
    }
}
