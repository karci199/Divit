#nullable enable
namespace Divit.Pdf;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

/// <summary>
/// A TrueType font with only the glyphs a document uses, for embedding. Glyph ids stay as they are, so text keeps
/// pointing at the same glyphs: unused glyphs are left empty, which costs four bytes each in the location table.
/// </summary>
internal static class FontSubset
{
    /// <summary>
    /// The tables a PDF reader needs to draw a TrueType font, in tag order; hinting tables only when present.
    /// </summary>
    private static readonly string[] Tags = { "OS/2", "cmap", "cvt ", "fpgm", "glyf", "head", "hhea", "hmtx", "loca", "maxp", "name", "post", "prep" };

    /// <summary>
    /// The subset font file; <paramref name="glyphs"/> grows by the parts of composite glyphs. Its character map holds
    /// the used characters only: readers draw by glyph id, but tools that extract or edit fonts expect one, as they
    /// expect the name table.
    /// </summary>
    public static byte[] Build(TrueTypeFont font, SortedSet<ushort> glyphs, IReadOnlyDictionary<ushort, int>? characters = null)
    {
        glyphs.Add(0);
        font.TryTable("loca", out var loca, out _);
        font.TryTable("glyf", out var glyf, out _);
        int Location(int glyph) => font.LongLoca ? (int)font.U32(loca + (4 * glyph)) : 2 * font.U16(loca + (2 * glyph));

        // Composite glyphs are made of others: take those too, until nothing new comes.
        var pending = new Stack<ushort>(glyphs);
        while (pending.Count > 0)
        {
            var glyph = pending.Pop();
            int start = Location(glyph), end = Location(glyph + 1);
            if (end - start < 10 || font.S16(glyf + start) >= 0)
            {
                continue;
            }

            var p = glyf + start + 10;
            while (true)
            {
                int flags = font.U16(p), part = font.U16(p + 2);
                if (part < font.GlyphCount && glyphs.Add((ushort)part))
                {
                    pending.Push((ushort)part);
                }

                p += 4 + ((flags & 0x0001) != 0 ? 4 : 2) + ((flags & 0x0008) != 0 ? 2 : (flags & 0x0040) != 0 ? 4 : (flags & 0x0080) != 0 ? 8 : 0);
                if ((flags & 0x0020) == 0)
                {
                    break;
                }
            }
        }

        // The outlines of the kept glyphs, each on a four-byte boundary, and their locations (long format).
        var outlines = new MemoryStream();
        var locations = new byte[4 * (font.GlyphCount + 1)];
        for (var g = 0; g < font.GlyphCount; g++)
        {
            Put32(locations, 4 * g, (uint)outlines.Length);
            if (glyphs.Contains((ushort)g))
            {
                int start = Location(g), end = Location(g + 1);
                if (end > start)
                {
                    outlines.Write(font.Data, glyf + start, end - start);
                    while (outlines.Length % 4 != 0)
                    {
                        outlines.WriteByte(0);
                    }
                }
            }
        }

        Put32(locations, 4 * font.GlyphCount, (uint)outlines.Length);

        var tables = new List<(string Tag, byte[] Data)>();
        foreach (var tag in Tags)
        {
            byte[] data;
            if (tag == "glyf")
            {
                data = outlines.ToArray();
            }
            else if (tag == "loca")
            {
                data = locations;
            }
            else if (tag == "cmap")
            {
                data = CharacterMap(characters);
            }
            else if (font.TryTable(tag, out var offset, out var length))
            {
                data = new byte[length];
                Buffer.BlockCopy(font.Data, offset, data, 0, length);
                if (tag == "head")
                {
                    // Long locations, and the checksum adjustment cleared until the whole file is summed.
                    data[50] = 0;
                    data[51] = 1;
                    Put32(data, 8, 0);
                }
                else if (tag == "post" && data.Length > 32)
                {
                    // Version 3: no glyph names, which only the full font needs.
                    data = Slice(data, 32);
                    Put32(data, 0, 0x00030000);
                }
            }
            else
            {
                continue;
            }

            tables.Add((tag, data));
        }

        return Assemble(tables);
    }

    /// <summary>
    /// A font file of the tables: the directory with checksums, each table on a four-byte boundary, and the
    /// head table's checksum adjustment over the whole file.
    /// </summary>
    private static byte[] Assemble(List<(string Tag, byte[] Data)> tables)
    {
        var count = tables.Count;
        var power = 1;
        var log = 0;
        while (power * 2 <= count)
        {
            power *= 2;
            log++;
        }

        var size = 12 + (16 * count);
        foreach (var table in tables)
        {
            size += (table.Data.Length + 3) & ~3;
        }

        var file = new byte[size];
        Put32(file, 0, 0x00010000);
        Put16(file, 4, count);
        Put16(file, 6, power * 16);
        Put16(file, 8, log);
        Put16(file, 10, (count * 16) - (power * 16));
        var offset = 12 + (16 * count);
        var head = -1;
        for (var i = 0; i < count; i++)
        {
            var (tag, data) = tables[i];
            var record = 12 + (16 * i);
            Encoding.ASCII.GetBytes(tag, 0, 4, file, record);
            Buffer.BlockCopy(data, 0, file, offset, data.Length);
            Put32(file, record + 4, Checksum(file, offset, data.Length));
            Put32(file, record + 8, (uint)offset);
            Put32(file, record + 12, (uint)data.Length);
            if (tag == "head")
            {
                head = offset;
            }

            offset += (data.Length + 3) & ~3;
        }

        if (head >= 0)
        {
            Put32(file, head + 8, 0xB1B0AFBA - Checksum(file, 0, file.Length));
        }

        return file;
    }

    /// <summary>
    /// A character map of the used characters of the basic plane: one Windows Unicode subtable (format 4), a
    /// segment per character, and the closing segment the format requires.
    /// </summary>
    private static byte[] CharacterMap(IReadOnlyDictionary<ushort, int>? characters)
    {
        var pairs = new SortedDictionary<int, ushort>();
        if (characters != null)
        {
            foreach (var pair in characters)
            {
                if (pair.Value < 0xFFFF && !pairs.ContainsKey(pair.Value))
                {
                    pairs[pair.Value] = pair.Key;
                }
            }
        }

        var segments = pairs.Count + 1;
        var power = 1;
        var log = 0;
        while (power * 2 <= segments)
        {
            power *= 2;
            log++;
        }

        var length = 16 + (8 * segments);
        var data = new byte[12 + length];
        Put16(data, 2, 1);
        Put16(data, 4, 3);
        Put16(data, 6, 1);
        Put32(data, 8, 12);
        var t = 12;
        Put16(data, t, 4);
        Put16(data, t + 2, length);
        Put16(data, t + 6, 2 * segments);
        Put16(data, t + 8, 2 * power);
        Put16(data, t + 10, log);
        Put16(data, t + 12, (2 * segments) - (2 * power));
        // The range offsets stay 0: each glyph comes from its delta.
        int ends = t + 14, starts = ends + (2 * segments) + 2, deltas = starts + (2 * segments);
        var s = 0;
        foreach (var pair in pairs)
        {
            Put16(data, ends + (2 * s), pair.Key);
            Put16(data, starts + (2 * s), pair.Key);
            Put16(data, deltas + (2 * s), (pair.Value - pair.Key) & 0xFFFF);
            s++;
        }

        Put16(data, ends + (2 * s), 0xFFFF);
        Put16(data, starts + (2 * s), 0xFFFF);
        Put16(data, deltas + (2 * s), 1);
        return data;
    }

    private static uint Checksum(byte[] data, int offset, int length)
    {
        uint sum = 0;
        var end = offset + ((length + 3) & ~3);
        for (var i = offset; i < end; i += 4)
        {
            uint word = 0;
            for (var k = 0; k < 4; k++)
            {
                word = (word << 8) | (i + k < offset + length ? data[i + k] : (byte)0);
            }

            sum += word;
        }

        return sum;
    }

    private static byte[] Slice(byte[] data, int length)
    {
        var slice = new byte[length];
        Buffer.BlockCopy(data, 0, slice, 0, length);
        return slice;
    }

    private static void Put16(byte[] data, int offset, int value)
    {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void Put32(byte[] data, int offset, uint value)
    {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
