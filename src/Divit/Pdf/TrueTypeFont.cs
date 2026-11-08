#nullable enable
namespace Divit.Pdf;

using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Text;

/// <summary>
/// A TrueType font read from its file: the tables, the metrics, the advance width of every glyph and the glyph of
/// every character. One font of a collection (.ttc) is picked by its index. Read once and shared: it never changes.
/// </summary>
internal sealed class TrueTypeFont
{
    private readonly Dictionary<string, (int Offset, int Length)> _tables = new Dictionary<string, (int, int)>(StringComparer.Ordinal);

    /// <summary>
    /// Glyph of each character of the basic plane, 0 when the font has none; others in <see cref="_supplementary"/>.
    /// </summary>
    private readonly ushort[] _glyphs = new ushort[65536];

    private readonly Dictionary<int, ushort> _supplementary = new Dictionary<int, ushort>();

    public TrueTypeFont(byte[] data, int index = 0)
    {
        Data = data;
        try
        {
            var start = 0;
            if (Tag(0) == "ttcf")
            {
                var count = (int)U32(8);
                start = (int)U32(12 + 4 * Math.Max(0, Math.Min(index, count - 1)));
            }

            var version = U32(start);
            if (version != 0x00010000 && Tag(start) != "true" && Tag(start) != "OTTO")
            {
                throw new ArgumentException(Messages.Get(Strings.FontInvalid));
            }

            int tables = U16(start + 4);
            for (var i = 0; i < tables; i++)
            {
                var record = start + 12 + (i * 16);
                _tables[Tag(record)] = ((int)U32(record + 8), (int)U32(record + 12));
            }

            if (_tables.TryGetValue("name", out var name))
            {
                (Family, Style, PostScriptName) = Names(Data, name.Offset);
            }

            if (_tables.ContainsKey("CFF ") || _tables.ContainsKey("CFF2") || !_tables.ContainsKey("glyf"))
            {
                throw new NotSupportedException(Messages.Format(Strings.FontOutlinesUnsupported, Family));
            }

            var head = Table("head");
            UnitsPerEm = U16(head + 18);
            XMin = S16(head + 36);
            YMin = S16(head + 38);
            XMax = S16(head + 40);
            YMax = S16(head + 42);
            LongLoca = S16(head + 50) == 1;
            var hhea = Table("hhea");
            Ascent = S16(hhea + 4);
            Descent = S16(hhea + 6);
            LineGap = S16(hhea + 8);
            int metrics = U16(hhea + 34);
            GlyphCount = U16(Table("maxp") + 4);
            Advances = new ushort[GlyphCount];
            var hmtx = Table("hmtx");
            for (var g = 0; g < GlyphCount; g++)
            {
                Advances[g] = (ushort)U16(hmtx + (4 * Math.Min(g, metrics - 1)));
            }

            CapHeight = Ascent;
            if (_tables.TryGetValue("OS/2", out var os2))
            {
                Weight = U16(os2.Offset + 4);
                var selection = U16(os2.Offset + 62);
                Bold = (selection & 0x20) != 0;
                Italic = (selection & 0x01) != 0;
                if (U16(os2.Offset) >= 2 && os2.Length >= 90)
                {
                    CapHeight = S16(os2.Offset + 88);
                }
            }

            if (_tables.TryGetValue("post", out var post))
            {
                ItalicAngle = S16(post.Offset + 4) + (U16(post.Offset + 6) / 65536.0);
                FixedPitch = U32(post.Offset + 12) != 0;
            }

            ReadCharacterMap();
        }
        catch (IndexOutOfRangeException)
        {
            throw new ArgumentException(Messages.Get(Strings.FontInvalid));
        }
    }

    public byte[] Data { get; }

    public string Family { get; private set; } = string.Empty;

    public string Style { get; private set; } = string.Empty;

    /// <summary>
    /// The PostScript name, without spaces: the font's name in the PDF.
    /// </summary>
    public string PostScriptName { get; private set; } = string.Empty;

    public bool Bold { get; private set; }

    public bool Italic { get; private set; }

    public int Weight { get; private set; } = 400;

    public bool FixedPitch { get; private set; }

    public double ItalicAngle { get; private set; }

    public int UnitsPerEm { get; }

    public int Ascent { get; }

    public int Descent { get; }

    public int LineGap { get; }

    public int CapHeight { get; }

    public int XMin { get; }

    public int YMin { get; }

    public int XMax { get; }

    public int YMax { get; }

    public int GlyphCount { get; }

    /// <summary>
    /// Advance width of each glyph in font units.
    /// </summary>
    public ushort[] Advances { get; }

    /// <summary>
    /// Whether the glyph locations are 32-bit offsets.
    /// </summary>
    public bool LongLoca { get; }

    /// <summary>
    /// The glyph of a character (a code point), 0 for the missing glyph.
    /// </summary>
    /// <summary>
    /// The kerning of the font, read when first needed.
    /// </summary>
    public FontKerning Kerning => _kerning ??= FontKerning.Read(this);

    private FontKerning? _kerning;

    public ushort Glyph(int codePoint) =>
        codePoint < 65536 ? _glyphs[codePoint] : _supplementary.TryGetValue(codePoint, out var glyph) ? glyph : (ushort)0;

    /// <summary>
    /// Offset of a table in <see cref="Data"/>, and whether the font has it.
    /// </summary>
    public bool TryTable(string tag, out int offset, out int length)
    {
        var found = _tables.TryGetValue(tag, out var table);
        offset = table.Offset;
        length = table.Length;
        return found;
    }

    public int U16(int offset) => (Data[offset] << 8) | Data[offset + 1];

    public int S16(int offset) => (short)U16(offset);

    public uint U32(int offset) => (uint)((Data[offset] << 24) | (Data[offset + 1] << 16) | (Data[offset + 2] << 8) | Data[offset + 3]);

    private int Table(string tag) => _tables.TryGetValue(tag, out var table) ? table.Offset : throw new ArgumentException(Messages.Get(Strings.FontInvalid));

    private string Tag(int offset) => Encoding.ASCII.GetString(Data, offset, 4);

    /// <summary>
    /// Family, style and PostScript names from the name table at <paramref name="table"/>, preferring the typographic
    /// names (ids 16, 17) and Windows English records.
    /// </summary>
    internal static (string Family, string Style, string PostScript) Names(byte[] data, int table)
    {
        int U16(int offset) => (data[offset] << 8) | data[offset + 1];
        int count = U16(table + 2), strings = table + U16(table + 4);
        var best = new Dictionary<int, (int Score, string Text)>();
        for (var i = 0; i < count; i++)
        {
            var record = table + 6 + (i * 12);
            int platform = U16(record), encoding = U16(record + 2), language = U16(record + 4), id = U16(record + 6);
            int length = U16(record + 8), offset = U16(record + 10);
            if (id != 1 && id != 2 && id != 4 && id != 6 && id != 16 && id != 17)
            {
                continue;
            }

            string text;
            int score;
            if (platform == 3 && (encoding == 1 || encoding == 0 || encoding == 10))
            {
                text = Encoding.BigEndianUnicode.GetString(data, strings + offset, length);
                score = language == 0x0409 ? 3 : 2;
            }
            else if (platform == 1 && encoding == 0)
            {
                text = Encoding.ASCII.GetString(data, strings + offset, length);
                score = 1;
            }
            else
            {
                continue;
            }

            if (!best.TryGetValue(id, out var current) || score > current.Score)
            {
                best[id] = (score, text);
            }
        }

        string Get(int id) => best.TryGetValue(id, out var entry) ? entry.Text : string.Empty;
        var family = Get(16).Length > 0 ? Get(16) : Get(1);
        var style = Get(17).Length > 0 ? Get(17) : Get(2);
        var postScript = Get(6).Length > 0 ? Get(6) : (family + style).Replace(" ", string.Empty);
        var clean = new StringBuilder();
        foreach (var c in postScript)
        {
            // A PDF name of plain characters: letters, digits, hyphens.
            if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '-')
            {
                clean.Append(c);
            }
        }

        return (family, style, clean.Length > 0 ? clean.ToString() : "Font");
    }

    /// <summary>
    /// The character map: the Unicode subtable with all planes (format 12) if there is one, else the basic plane
    /// (format 4).
    /// </summary>
    private void ReadCharacterMap()
    {
        if (!_tables.TryGetValue("cmap", out var map))
        {
            // A subset for embedding has none: text is written as glyph ids.
            return;
        }

        var cmap = map.Offset;
        int count = U16(cmap + 2);
        int format4 = -1, format12 = -1;
        for (var i = 0; i < count; i++)
        {
            var record = cmap + 4 + (i * 8);
            int platform = U16(record), encoding = U16(record + 2);
            var sub = cmap + (int)U32(record + 4);
            var format = U16(sub);
            if (format == 12 && ((platform == 3 && encoding == 10) || platform == 0))
            {
                format12 = sub;
            }
            else if (format == 4 && ((platform == 3 && (encoding == 1 || encoding == 0)) || platform == 0) && format4 < 0)
            {
                format4 = sub;
            }
        }

        if (format12 >= 0)
        {
            var groups = (int)U32(format12 + 12);
            for (var i = 0; i < groups; i++)
            {
                var group = format12 + 16 + (i * 12);
                long first = U32(group), last = U32(group + 4), glyph = U32(group + 8);
                for (var c = first; c <= last && c <= 0x10FFFF; c++)
                {
                    Map((int)c, (ushort)(glyph + c - first));
                }
            }

            return;
        }

        if (format4 < 0)
        {
            return;
        }

        var segments = U16(format4 + 6) / 2;
        int ends = format4 + 14, starts = ends + (segments * 2) + 2, deltas = starts + (segments * 2), ranges = deltas + (segments * 2);
        for (var s = 0; s < segments; s++)
        {
            int end = U16(ends + (s * 2)), first = U16(starts + (s * 2)), delta = S16(deltas + (s * 2)), range = U16(ranges + (s * 2));
            for (var c = first; c <= end && c != 0xFFFF; c++)
            {
                int glyph;
                if (range == 0)
                {
                    glyph = (c + delta) & 0xFFFF;
                }
                else
                {
                    glyph = U16(ranges + (s * 2) + range + ((c - first) * 2));
                    if (glyph != 0)
                    {
                        glyph = (glyph + delta) & 0xFFFF;
                    }
                }

                Map(c, (ushort)glyph);
            }
        }
    }

    private void Map(int codePoint, ushort glyph)
    {
        if (glyph >= GlyphCount)
        {
            return;
        }

        if (codePoint < 65536)
        {
            _glyphs[codePoint] = glyph;
        }
        else
        {
            _supplementary[codePoint] = glyph;
        }
    }
}
