#nullable enable
namespace Divit.Xlsx;

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

/// <content>
/// The shared string table.
/// </content>
public sealed partial class XlsxWorkbook
{
    /// <summary>
    /// Shared strings: each distinct text once, in order of first use.
    /// </summary>
    private sealed class StringTable
    {
        private readonly Dictionary<string, int> _index = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<string> _texts = new List<string>();
        private int _uses;

        public int Index(string text)
        {
            _uses++;
            if (!_index.TryGetValue(text, out var i))
            {
                i = _texts.Count;
                _texts.Add(text);
                _index[text] = i;
            }

            return i;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("sst", MainNs);
            w.WriteAttributeString("count", Int(_uses));
            w.WriteAttributeString("uniqueCount", Int(_texts.Count));
            foreach (var text in _texts)
            {
                w.WriteStartElement("si", MainNs);
                w.WriteStartElement("t", MainNs);
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]) || text.IndexOf('\n') >= 0))
                {
                    w.WriteAttributeString("xml", "space", null, "preserve");
                }

                w.WriteString(Clean(text));
                w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        /// <summary>
        /// Drops control characters that XML cannot hold (tab and line feed stay); XmlWriter throws on them.
        /// </summary>
        public static string Clean(string text)
        {
            StringBuilder? sb = null;
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                var bad = ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r';
                if (bad && sb == null)
                {
                    sb = new StringBuilder(text, 0, i, text.Length);
                }
                else if (!bad)
                {
                    sb?.Append(ch);
                }
            }

            return sb?.ToString() ?? text;
        }
    }
}
