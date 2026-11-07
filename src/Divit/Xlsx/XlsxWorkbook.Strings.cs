#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using System;
using System.Collections.Generic;
using System.Xml;

/// <content>
/// The shared string table.
/// </content>
public sealed partial class XlsxWorkbook
{
    /// <summary>
    /// Shared strings: each distinct text once, in order of first use. Plain texts and rich texts share the numbering
    /// but have their own lookups, so plain texts keep the fast ordinal string comparison.
    /// </summary>
    private sealed class StringTable
    {
        private readonly XlsxWorkbook _book;
        private readonly Dictionary<string, int> _index = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly Dictionary<XlsxRichText, int> _richIndex = new Dictionary<XlsxRichText, int>();
        private readonly List<object> _items = new List<object>();
        private int _uses;

        public StringTable(XlsxWorkbook book)
        {
            _book = book;
        }

        public int Index(string text)
        {
            _uses++;
            if (!_index.TryGetValue(text, out var i))
            {
                i = _items.Count;
                _items.Add(text);
                _index[text] = i;
            }

            return i;
        }

        public int Index(XlsxRichText text)
        {
            _uses++;
            if (!_richIndex.TryGetValue(text, out var i))
            {
                i = _items.Count;
                _items.Add(text);
                _richIndex[text] = i;
            }

            return i;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("sst", MainNs);
            w.WriteAttributeString("count", Int(_uses));
            w.WriteAttributeString("uniqueCount", Int(_items.Count));
            foreach (var item in _items)
            {
                w.WriteStartElement("si", MainNs);
                if (item is string text)
                {
                    Text(w, text);
                }
                else
                {
                    Runs(w, (XlsxRichText)item);
                }

                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        /// <summary>
        /// The pieces of a rich text, inside a shared string item or an inline string.
        /// </summary>
        public void Runs(XmlWriter w, XlsxRichText text)
        {
            foreach (var run in text.Runs)
            {
                w.WriteStartElement("r", MainNs);
                if (run.Look is XlsxStyle look)
                {
                    RunProperties(w, look);
                }

                Text(w, run.Text);
                w.WriteEndElement();
            }
        }

        /// <summary>
        /// A <c>t</c> element; spaces at the ends and line breaks are kept.
        /// </summary>
        public static void Text(XmlWriter w, string text)
        {
            w.WriteStartElement("t", MainNs);
            if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]) || text.IndexOf('\n') >= 0))
            {
                w.WriteAttributeString("xml", "space", null, "preserve");
            }

            w.WriteString(OfficeText.Clean(text));
            w.WriteEndElement();
        }

        /// <summary>
        /// The font of a piece, complete: a piece with properties takes nothing from the cell, so size and font name
        /// fall back to the workbook's. Children in the order Excel writes them.
        /// </summary>
        private void RunProperties(XmlWriter w, XlsxStyle s)
        {
            w.WriteStartElement("rPr", MainNs);
            if (s.Bold) { w.WriteElementString("b", MainNs, null); }
            if (s.Italic) { w.WriteElementString("i", MainNs, null); }
            if (s.Strikethrough) { w.WriteElementString("strike", MainNs, null); }
            if (s.DoubleUnderline) { w.WriteStartElement("u", MainNs); w.WriteAttributeString("val", "double"); w.WriteEndElement(); }
            else if (s.Underline) { w.WriteElementString("u", MainNs, null); }
            if (s.Script != XlsxScript.None)
            {
                w.WriteStartElement("vertAlign", MainNs);
                w.WriteAttributeString("val", s.Script == XlsxScript.Superscript ? "superscript" : "subscript");
                w.WriteEndElement();
            }

            w.WriteStartElement("sz", MainNs);
            w.WriteAttributeString("val", Num(s.FontSize > 0 ? s.FontSize : _book.DefaultFontSize));
            w.WriteEndElement();
            if (s.FontColor != null)
            {
                w.WriteStartElement("color", MainNs);
                w.WriteAttributeString("rgb", Argb(s.FontColor));
                w.WriteEndElement();
            }

            w.WriteStartElement("rFont", MainNs);
            w.WriteAttributeString("val", s.FontName ?? _book.DefaultFontName);
            w.WriteEndElement();
            w.WriteEndElement();
        }
    }
}
