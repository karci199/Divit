#nullable enable
namespace Divit.Pdf;

using System;
using System.Collections.Generic;

/// <summary>
/// The kerning of a font: how much closer (or further) two glyphs stand side by side, such as "AV" or "To". Read
/// from the pair adjustments of the GPOS table's kern feature, or from the old kern table of fonts without one.
/// Values are in font units and added to the first glyph's advance.
/// </summary>
internal sealed class FontKerning
{
    /// <summary>
    /// No kerning: for fonts without either table.
    /// </summary>
    public static readonly FontKerning None = new FontKerning();

    private readonly Dictionary<uint, short> _pairs = new Dictionary<uint, short>();
    private readonly List<List<Subtable>> _lookups = new List<List<Subtable>>();

    private FontKerning()
    {
    }

    /// <summary>
    /// Whether the font kerns at all.
    /// </summary>
    public bool Any => _pairs.Count > 0 || _lookups.Count > 0;

    /// <summary>
    /// Reads the kerning of a font: GPOS when it has a kern feature, else the kern table.
    /// </summary>
    public static FontKerning Read(TrueTypeFont font, bool gpos = true)
    {
        var kerning = new FontKerning();
        try
        {
            if (gpos && font.TryTable("GPOS", out var table, out _))
            {
                kerning.ReadPositioning(font, table);
            }

            if (!kerning.Any && font.TryTable("kern", out var kern, out _))
            {
                kerning.ReadKernTable(font, kern);
            }
        }
        catch (IndexOutOfRangeException)
        {
            // A damaged table: the font draws without kerning.
            return None;
        }

        return kerning.Any ? kerning : None;
    }

    /// <summary>
    /// The adjustment between two glyphs in font units; 0 when they are not kerned.
    /// </summary>
    public int Pair(ushort left, ushort right)
    {
        if (_pairs.Count > 0 && _pairs.TryGetValue(((uint)left << 16) | right, out var value))
        {
            return value;
        }

        var total = 0;
        foreach (var lookup in _lookups)
        {
            foreach (var subtable in lookup)
            {
                if (subtable.Adjust(left, right, out var adjustment))
                {
                    total += adjustment;
                    break;
                }
            }
        }

        return total;
    }

    /// <summary>
    /// The lookups of the kern features (any script): pair adjustments (type 2), also behind extensions (type 9).
    /// </summary>
    private void ReadPositioning(TrueTypeFont f, int gpos)
    {
        int features = gpos + f.U16(gpos + 6), lookups = gpos + f.U16(gpos + 8);
        var indices = new SortedSet<int>();
        int featureCount = f.U16(features);
        for (var i = 0; i < featureCount; i++)
        {
            var record = features + 2 + (i * 6);
            if (f.Data[record] == 'k' && f.Data[record + 1] == 'e' && f.Data[record + 2] == 'r' && f.Data[record + 3] == 'n')
            {
                var feature = features + f.U16(record + 4);
                int count = f.U16(feature + 2);
                for (var k = 0; k < count; k++)
                {
                    indices.Add(f.U16(feature + 4 + (2 * k)));
                }
            }
        }

        int lookupCount = f.U16(lookups);
        foreach (var index in indices)
        {
            if (index >= lookupCount)
            {
                continue;
            }

            var lookup = lookups + f.U16(lookups + 2 + (2 * index));
            int type = f.U16(lookup), subtables = f.U16(lookup + 4);
            var list = new List<Subtable>();
            for (var s = 0; s < subtables; s++)
            {
                var subtable = lookup + f.U16(lookup + 6 + (2 * s));
                var kind = type;
                if (type == 9)
                {
                    kind = f.U16(subtable + 2);
                    subtable += (int)f.U32(subtable + 4);
                }

                if (kind == 2)
                {
                    list.Add(new Subtable(f, subtable));
                }
            }

            if (list.Count > 0)
            {
                _lookups.Add(list);
            }
        }
    }

    /// <summary>
    /// The old kern table: horizontal pair lists (format 0).
    /// </summary>
    private void ReadKernTable(TrueTypeFont f, int kern)
    {
        int tables = f.U16(kern + 2);
        var p = kern + 4;
        for (var t = 0; t < tables; t++)
        {
            int length = f.U16(p + 2), coverage = f.U16(p + 4);
            if ((coverage >> 8) == 0 && (coverage & 1) != 0)
            {
                int pairs = f.U16(p + 6);
                for (var i = 0; i < pairs; i++)
                {
                    var record = p + 14 + (6 * i);
                    var key = ((uint)f.U16(record) << 16) | (uint)f.U16(record + 2);
                    if (!_pairs.ContainsKey(key))
                    {
                        _pairs[key] = (short)f.S16(record + 4);
                    }
                }
            }

            p += length;
        }
    }

    /// <summary>
    /// One pair adjustment subtable: pairs listed glyph by glyph (format 1), or by glyph classes (format 2).
    /// </summary>
    private sealed class Subtable
    {
        private readonly bool[] _covered;
        private readonly Dictionary<uint, short>? _pairs;
        private readonly ushort[]? _first;
        private readonly ushort[]? _second;
        private readonly short[]? _matrix;
        private readonly int _columns;

        public Subtable(TrueTypeFont f, int table)
        {
            int format = f.U16(table), valueFormat1 = f.U16(table + 4), valueFormat2 = f.U16(table + 6);
            var coverage = Coverage(f, table + f.U16(table + 2));
            _covered = new bool[f.GlyphCount];
            foreach (var glyph in coverage)
            {
                if (glyph < _covered.Length)
                {
                    _covered[glyph] = true;
                }
            }

            int size1 = Size(valueFormat1), size2 = Size(valueFormat2);
            var advance = (valueFormat1 & 0x0004) != 0 ? 2 * Bits(valueFormat1 & 0x0003) : -1;
            if (format == 1)
            {
                _pairs = new Dictionary<uint, short>();
                int sets = f.U16(table + 8);
                for (var i = 0; i < sets && i < coverage.Count; i++)
                {
                    var set = table + f.U16(table + 10 + (2 * i));
                    int count = f.U16(set);
                    var record = set + 2;
                    for (var k = 0; k < count; k++, record += 2 + size1 + size2)
                    {
                        if (advance >= 0)
                        {
                            _pairs[((uint)coverage[i] << 16) | (uint)f.U16(record)] = (short)f.S16(record + 2 + advance);
                        }
                    }
                }
            }
            else if (format == 2)
            {
                _first = Classes(f, table + f.U16(table + 8));
                _second = Classes(f, table + f.U16(table + 10));
                int rows = f.U16(table + 12);
                _columns = f.U16(table + 14);
                _matrix = new short[rows * _columns];
                var record = table + 16;
                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < _columns; c++, record += size1 + size2)
                    {
                        _matrix[(r * _columns) + c] = advance >= 0 ? (short)f.S16(record + advance) : (short)0;
                    }
                }
            }
        }

        /// <summary>
        /// The adjustment of a pair, when the subtable covers the first glyph: format 1 only for listed pairs.
        /// </summary>
        public bool Adjust(ushort left, ushort right, out int adjustment)
        {
            adjustment = 0;
            if (left >= _covered.Length || !_covered[left])
            {
                return false;
            }

            if (_pairs != null)
            {
                if (!_pairs.TryGetValue(((uint)left << 16) | right, out var value))
                {
                    return false;
                }

                adjustment = value;
                return true;
            }

            if (_matrix == null)
            {
                return false;
            }

            int row = left < _first!.Length ? _first[left] : 0, column = right < _second!.Length ? _second[right] : 0;
            var index = (row * _columns) + column;
            adjustment = index < _matrix.Length ? _matrix[index] : 0;
            return true;
        }

        private static List<ushort> Coverage(TrueTypeFont f, int table)
        {
            var glyphs = new List<ushort>();
            int format = f.U16(table), count = f.U16(table + 2);
            for (var i = 0; i < count; i++)
            {
                if (format == 1)
                {
                    glyphs.Add((ushort)f.U16(table + 4 + (2 * i)));
                }
                else
                {
                    var range = table + 4 + (6 * i);
                    for (int g = f.U16(range), end = f.U16(range + 2); g <= end; g++)
                    {
                        glyphs.Add((ushort)g);
                    }
                }
            }

            return glyphs;
        }

        private static ushort[] Classes(TrueTypeFont f, int table)
        {
            var classes = new ushort[f.GlyphCount];
            int format = f.U16(table);
            if (format == 1)
            {
                int start = f.U16(table + 2), count = f.U16(table + 4);
                for (var i = 0; i < count && start + i < classes.Length; i++)
                {
                    classes[start + i] = (ushort)f.U16(table + 6 + (2 * i));
                }
            }
            else if (format == 2)
            {
                int ranges = f.U16(table + 2);
                for (var i = 0; i < ranges; i++)
                {
                    var range = table + 4 + (6 * i);
                    for (int g = f.U16(range), end = f.U16(range + 2), value = f.U16(range + 4); g <= end && g < classes.Length; g++)
                    {
                        classes[g] = (ushort)value;
                    }
                }
            }

            return classes;
        }

        /// <summary>
        /// Bytes of a value record: two per field the format has.
        /// </summary>
        private static int Size(int format) => 2 * Bits(format & 0xFF);

        private static int Bits(int value)
        {
            var count = 0;
            for (; value != 0; value &= value - 1)
            {
                count++;
            }

            return count;
        }
    }
}
