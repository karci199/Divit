#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

/// <content>
/// The style table (styles.xml).
/// </content>
public sealed partial class XlsxWorkbook
{
    /// <summary>
    /// Style table: each distinct <see cref="XlsxStyle"/> is one <c>xf</c> record; fonts, fills, borders and number
    /// formats are made unique separately. Record 0 is the default style.
    /// </summary>
    private sealed class StyleTable
    {
        private readonly XlsxWorkbook _book;
        private readonly Dictionary<XlsxStyle, int> _xfIndex = new Dictionary<XlsxStyle, int>();
        private readonly List<(int Font, int Fill, int Border, int NumFmt, XlsxStyle Style)> _xfs = new List<(int, int, int, int, XlsxStyle)>();
        private readonly List<string> _fonts = new List<string>();
        private readonly List<string> _fills = new List<string>();
        private readonly List<string> _borders = new List<string>();
        private readonly List<string> _numFmts = new List<string>();

        /// <summary>
        /// Differential formats of conditional rules, each distinct look once.
        /// </summary>
        private readonly List<XlsxStyle> _dxfs = new List<XlsxStyle>();

        /// <summary>
        /// The style asked for last and its record; starts as the default style, record 0.
        /// </summary>
        private XlsxStyle _last;

        private int _lastIndex;

        public StyleTable(XlsxWorkbook book)
        {
            _book = book;

            // Required start of the standard: font 0 is the default, fills 0 and 1 are none and gray125, border 0 is empty.
            _fonts.Add(FontKey(default(XlsxStyle)));
            _fills.Add("none");
            _fills.Add("gray125");
            _borders.Add(BorderKey(default(XlsxStyle)));
            _xfIndex[default(XlsxStyle)] = 0;
            _xfs.Add((0, 0, 0, 0, default(XlsxStyle)));
        }

        public int Index(XlsxStyle style)
        {
            // Neighbouring cells usually share a style; comparing with the last one is cheaper than hashing every string.
            if (style.Equals(_last))
            {
                return _lastIndex;
            }

            _last = style;
            if (_xfIndex.TryGetValue(style, out var i))
            {
                _lastIndex = i;
                return i;
            }

            var font = Add(_fonts, FontKey(style));
            var fill = style.Fill == null ? 0 : Add(_fills, "solid:" + Argb(style.Fill));
            var border = Add(_borders, BorderKey(style));
            var numFmt = style.NumberFormat == null ? 0 : 164 + Add(_numFmts, style.NumberFormat);
            i = _xfs.Count;
            _xfs.Add((font, fill, border, numFmt, style));
            _xfIndex[style] = i;
            _lastIndex = i;
            return i;
        }

        /// <summary>
        /// The record of a conditional look; its number format joins the workbook's number formats.
        /// </summary>
        public int Differential(XlsxStyle look)
        {
            var i = _dxfs.IndexOf(look);
            if (i < 0)
            {
                i = _dxfs.Count;
                _dxfs.Add(look);
                if (look.NumberFormat != null)
                {
                    Add(_numFmts, look.NumberFormat);
                }
            }

            return i;
        }

        public void Write(XmlWriter w)
        {
            w.WriteStartElement("styleSheet", MainNs);

            if (_numFmts.Count > 0)
            {
                w.WriteStartElement("numFmts", MainNs);
                w.WriteAttributeString("count", Int(_numFmts.Count));
                for (var i = 0; i < _numFmts.Count; i++)
                {
                    w.WriteStartElement("numFmt", MainNs);
                    w.WriteAttributeString("numFmtId", Int(164 + i));
                    w.WriteAttributeString("formatCode", _numFmts[i]);
                    w.WriteEndElement();
                }

                w.WriteEndElement();
            }

            // Fonts: key "name|size|bold|italic|underline|colour|strike|script"; underline 1 single, 2 double; script 1
            // superscript, 2 subscript. Children in the order Excel writes them.
            w.WriteStartElement("fonts", MainNs);
            w.WriteAttributeString("count", Int(_fonts.Count));
            foreach (var key in _fonts)
            {
                var p = key.Split('|');
                w.WriteStartElement("font", MainNs);
                if (p[2] == "1") { w.WriteElementString("b", MainNs, null); }
                if (p[3] == "1") { w.WriteElementString("i", MainNs, null); }
                if (p[6] == "1") { w.WriteElementString("strike", MainNs, null); }
                if (p[4] == "1") { w.WriteElementString("u", MainNs, null); }
                if (p[4] == "2") { w.WriteStartElement("u", MainNs); w.WriteAttributeString("val", "double"); w.WriteEndElement(); }
                if (p[7] != "0")
                {
                    w.WriteStartElement("vertAlign", MainNs);
                    w.WriteAttributeString("val", p[7] == "1" ? "superscript" : "subscript");
                    w.WriteEndElement();
                }

                w.WriteStartElement("sz", MainNs); w.WriteAttributeString("val", p[1]); w.WriteEndElement();
                if (p[5].Length > 0) { w.WriteStartElement("color", MainNs); w.WriteAttributeString("rgb", p[5]); w.WriteEndElement(); }
                w.WriteStartElement("name", MainNs); w.WriteAttributeString("val", p[0]); w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();

            w.WriteStartElement("fills", MainNs);
            w.WriteAttributeString("count", Int(_fills.Count));
            foreach (var key in _fills)
            {
                w.WriteStartElement("fill", MainNs);
                w.WriteStartElement("patternFill", MainNs);
                if (key.StartsWith("solid:", StringComparison.Ordinal))
                {
                    w.WriteAttributeString("patternType", "solid");
                    w.WriteStartElement("fgColor", MainNs); w.WriteAttributeString("rgb", key.Substring(6)); w.WriteEndElement();
                    w.WriteStartElement("bgColor", MainNs); w.WriteAttributeString("indexed", "64"); w.WriteEndElement();
                }
                else
                {
                    w.WriteAttributeString("patternType", key);
                }

                w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();

            // Borders: key "left|right|top|bottom", each "kind,colour".
            w.WriteStartElement("borders", MainNs);
            w.WriteAttributeString("count", Int(_borders.Count));
            foreach (var key in _borders)
            {
                var p = key.Split('|');
                w.WriteStartElement("border", MainNs);
                Side(w, "left", p[0]);
                Side(w, "right", p[1]);
                Side(w, "top", p[2]);
                Side(w, "bottom", p[3]);
                w.WriteStartElement("diagonal", MainNs); w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();

            w.WriteStartElement("cellStyleXfs", MainNs);
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("xf", MainNs);
            w.WriteAttributeString("numFmtId", "0"); w.WriteAttributeString("fontId", "0"); w.WriteAttributeString("fillId", "0"); w.WriteAttributeString("borderId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            w.WriteStartElement("cellXfs", MainNs);
            w.WriteAttributeString("count", Int(_xfs.Count));
            foreach (var xf in _xfs)
            {
                var s = xf.Style;
                w.WriteStartElement("xf", MainNs);
                w.WriteAttributeString("numFmtId", Int(xf.NumFmt));
                w.WriteAttributeString("fontId", Int(xf.Font));
                w.WriteAttributeString("fillId", Int(xf.Fill));
                w.WriteAttributeString("borderId", Int(xf.Border));
                w.WriteAttributeString("xfId", "0");
                if (xf.NumFmt != 0) { w.WriteAttributeString("applyNumberFormat", "1"); }
                if (xf.Font != 0) { w.WriteAttributeString("applyFont", "1"); }
                if (xf.Fill != 0) { w.WriteAttributeString("applyFill", "1"); }
                if (xf.Border != 0) { w.WriteAttributeString("applyBorder", "1"); }
                var aligned = s.Horizontal != null || s.Vertical != null || s.Wrap || s.ShrinkToFit || s.Indent > 0 || s.Rotation != 0;
                if (aligned)
                {
                    w.WriteAttributeString("applyAlignment", "1");
                    w.WriteStartElement("alignment", MainNs);

                    // An indent needs an edge to count from; Excel itself picks left.
                    var horizontal = s.Horizontal ?? (s.Indent > 0 ? "left" : null);
                    if (horizontal != null) { w.WriteAttributeString("horizontal", horizontal); }
                    if (s.Vertical != null) { w.WriteAttributeString("vertical", s.Vertical); }
                    if (s.Rotation != 0) { w.WriteAttributeString("textRotation", Int(TextRotation(s.Rotation))); }
                    if (s.Wrap) { w.WriteAttributeString("wrapText", "1"); }
                    if (s.Indent > 0) { w.WriteAttributeString("indent", Int(Math.Min(s.Indent, 250))); }
                    if (s.ShrinkToFit) { w.WriteAttributeString("shrinkToFit", "1"); }
                    w.WriteEndElement();
                }

                w.WriteEndElement();
            }

            w.WriteEndElement();

            w.WriteStartElement("cellStyles", MainNs);
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("cellStyle", MainNs);
            w.WriteAttributeString("name", "Normal"); w.WriteAttributeString("xfId", "0"); w.WriteAttributeString("builtinId", "0");
            w.WriteEndElement();
            w.WriteEndElement();

            if (_dxfs.Count > 0)
            {
                w.WriteStartElement("dxfs", MainNs);
                w.WriteAttributeString("count", Int(_dxfs.Count));
                foreach (var look in _dxfs)
                {
                    WriteDifferential(w, look);
                }

                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        /// <summary>
        /// One differential format: only what the look sets, in the order font, number format, fill, border.
        /// </summary>
        private void WriteDifferential(XmlWriter w, XlsxStyle s)
        {
            w.WriteStartElement("dxf", MainNs);
            if (s.Bold || s.Italic || s.Strikethrough || s.Underline || s.DoubleUnderline || s.FontColor != null)
            {
                w.WriteStartElement("font", MainNs);
                if (s.Bold) { w.WriteElementString("b", MainNs, null); }
                if (s.Italic) { w.WriteElementString("i", MainNs, null); }
                if (s.Strikethrough) { w.WriteElementString("strike", MainNs, null); }
                if (s.DoubleUnderline) { w.WriteStartElement("u", MainNs); w.WriteAttributeString("val", "double"); w.WriteEndElement(); }
                else if (s.Underline) { w.WriteElementString("u", MainNs, null); }
                if (s.FontColor != null) { w.WriteStartElement("color", MainNs); w.WriteAttributeString("rgb", Argb(s.FontColor)); w.WriteEndElement(); }
                w.WriteEndElement();
            }

            if (s.NumberFormat != null)
            {
                w.WriteStartElement("numFmt", MainNs);
                w.WriteAttributeString("numFmtId", Int(164 + _numFmts.IndexOf(s.NumberFormat)));
                w.WriteAttributeString("formatCode", s.NumberFormat);
                w.WriteEndElement();
            }

            // A differential solid fill takes its colour from bgColor, unlike a cell fill.
            if (s.Fill != null)
            {
                w.WriteStartElement("fill", MainNs);
                w.WriteStartElement("patternFill", MainNs);
                w.WriteStartElement("bgColor", MainNs);
                w.WriteAttributeString("rgb", Argb(s.Fill));
                w.WriteEndElement();
                w.WriteEndElement();
                w.WriteEndElement();
            }

            var border = BorderKey(s).Split('|');
            if (border.Any(p => p.Length > 0))
            {
                w.WriteStartElement("border", MainNs);
                var names = new[] { "left", "right", "top", "bottom" };
                for (var i = 0; i < 4; i++)
                {
                    if (border[i].Length > 0) { Side(w, names[i], border[i]); }
                }

                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        private string FontKey(XlsxStyle s) =>
            (s.FontName ?? _book.DefaultFontName) + "|" + Num(s.FontSize > 0 ? s.FontSize : _book.DefaultFontSize) + "|"
            + (s.Bold ? "1" : "0") + "|" + (s.Italic ? "1" : "0") + "|" + (s.DoubleUnderline ? "2" : s.Underline ? "1" : "0") + "|"
            + (s.FontColor == null ? string.Empty : Argb(s.FontColor)) + "|" + (s.Strikethrough ? "1" : "0") + "|"
            + Int((int)s.Script);

        /// <summary>
        /// Degrees to Excel's textRotation: 0–90 counter-clockwise, 91–180 for 1–90 clockwise, 255 stacked.
        /// </summary>
        private static int TextRotation(int degrees)
        {
            if (degrees == XlsxStyle.VerticalText)
            {
                return degrees;
            }

            degrees = Math.Max(-90, Math.Min(90, degrees));
            return degrees >= 0 ? degrees : 90 - degrees;
        }

        private static string BorderKey(XlsxStyle s) => Edge(s.Left) + "|" + Edge(s.Right) + "|" + Edge(s.Top) + "|" + Edge(s.Bottom);

        private static string Edge(XlsxBorder b) => b.Line == null ? string.Empty : b.Line + "," + (b.Color == null ? string.Empty : Argb(b.Color));

        private static void Side(XmlWriter w, string name, string key)
        {
            w.WriteStartElement(name, MainNs);
            if (key.Length > 0)
            {
                var p = key.Split(',');
                w.WriteAttributeString("style", p[0]);
                if (p[1].Length > 0)
                {
                    w.WriteStartElement("color", MainNs);
                    w.WriteAttributeString("rgb", p[1]);
                    w.WriteEndElement();
                }
            }

            w.WriteEndElement();
        }

        private static int Add(List<string> list, string key)
        {
            var i = list.IndexOf(key);
            if (i < 0)
            {
                i = list.Count;
                list.Add(key);
            }

            return i;
        }
    }
}
