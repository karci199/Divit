#nullable enable
namespace Divit.Pdf;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

/// <summary>
/// Finds installed fonts by family and style. The font folders of Windows, Linux and macOS are indexed once, from
/// the names and style bits of each file only; read fonts are kept, since they never change.
/// </summary>
internal static class FontLocator
{
    private static readonly object Gate = new object();
    private static Dictionary<string, List<Entry>>? _index;
    private static readonly Dictionary<(string Path, int Index), TrueTypeFont> Loaded = new Dictionary<(string, int), TrueTypeFont>();

    /// <summary>
    /// The installed font of a family closest to the style asked for; <c>null</c> when the family is not installed.
    /// </summary>
    public static TrueTypeFont? Find(string family, bool bold, bool italic)
    {
        Entry? best = null;
        lock (Gate)
        {
            _index ??= BuildIndex();
            if (!_index.TryGetValue(family.Trim().ToUpperInvariant(), out var entries))
            {
                return null;
            }

            var bestScore = int.MaxValue;
            foreach (var entry in entries)
            {
                // A wrong weight costs more than a wrong slant; a style name that is not plain costs a little.
                var score = (entry.Bold != bold ? 4 : 0) + (entry.Italic != italic ? 2 : 0) + (entry.Plain ? 0 : 1);
                if (score < bestScore)
                {
                    bestScore = score;
                    best = entry;
                }
            }
        }

        return best == null ? null : Load(best.Path, best.Index);
    }

    /// <summary>
    /// A font file read once and kept.
    /// </summary>
    public static TrueTypeFont Load(string path, int index)
    {
        lock (Gate)
        {
            if (!Loaded.TryGetValue((path, index), out var font))
            {
                font = new TrueTypeFont(File.ReadAllBytes(path), index);
                Loaded[(path, index)] = font;
            }

            return font;
        }
    }

    private static Dictionary<string, List<Entry>> BuildIndex()
    {
        var index = new Dictionary<string, List<Entry>>(StringComparer.Ordinal);
        foreach (var folder in Folders())
        {
            if (!Directory.Exists(folder))
            {
                continue;
            }

            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories);
            }
            catch (UnauthorizedAccessException)
            {
                continue;
            }

            foreach (var file in files)
            {
                var extension = Path.GetExtension(file).ToUpperInvariant();
                if (extension != ".TTF" && extension != ".TTC" && extension != ".OTF")
                {
                    continue;
                }

                try
                {
                    Describe(file, index);
                }
                catch (IOException)
                {
                }
                catch (UnauthorizedAccessException)
                {
                }
                catch (ArgumentException)
                {
                }
                catch (IndexOutOfRangeException)
                {
                }
            }
        }

        return index;
    }

    private static IEnumerable<string> Folders()
    {
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        if (windows.Length > 0)
        {
            yield return windows;
        }

        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (local.Length > 0)
        {
            yield return Path.Combine(local, "Microsoft", "Windows", "Fonts");
        }

        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        foreach (var folder in new[] { "/usr/share/fonts", "/usr/local/share/fonts", "/Library/Fonts", "/System/Library/Fonts" })
        {
            yield return folder;
        }

        if (home.Length > 0)
        {
            yield return Path.Combine(home, ".fonts");
            yield return Path.Combine(home, ".local", "share", "fonts");
            yield return Path.Combine(home, "Library", "Fonts");
        }
    }

    /// <summary>
    /// Adds the fonts of a file to the index, reading only its table directories, name tables and style bits.
    /// </summary>
    private static void Describe(string path, Dictionary<string, List<Entry>> index)
    {
        using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096))
        {
            var header = Read(file, 0, 12);
            var starts = new List<long>();
            if (Encoding.ASCII.GetString(header, 0, 4) == "ttcf")
            {
                var count = BigEndian(header, 8);
                var offsets = Read(file, 12, 4 * count);
                for (var i = 0; i < count; i++)
                {
                    starts.Add(BigEndian(offsets, 4 * i));
                }
            }
            else
            {
                starts.Add(0);
            }

            for (var i = 0; i < starts.Count; i++)
            {
                var directory = Read(file, starts[i], 12);
                var tables = (directory[4] << 8) | directory[5];
                var records = Read(file, starts[i] + 12, 16 * tables);
                long name = -1, os2 = -1;
                int nameLength = 0;
                var outlines = false;
                for (var t = 0; t < tables; t++)
                {
                    var tag = Encoding.ASCII.GetString(records, 16 * t, 4);
                    long offset = (uint)BigEndian(records, (16 * t) + 8);
                    if (tag == "name") { name = offset; nameLength = BigEndian(records, (16 * t) + 12); }
                    if (tag == "OS/2") { os2 = offset; }
                    if (tag == "glyf") { outlines = true; }
                }

                if (name < 0 || !outlines)
                {
                    continue;
                }

                var (family, style, _) = TrueTypeFont.Names(Read(file, name, nameLength), 0);
                bool bold = false, italic = false;
                if (os2 >= 0)
                {
                    var bits = Read(file, os2 + 62, 2);
                    bold = (bits[1] & 0x20) != 0;
                    italic = (bits[1] & 0x01) != 0;
                }

                var key = family.Trim().ToUpperInvariant();
                if (!index.TryGetValue(key, out var entries))
                {
                    entries = new List<Entry>();
                    index[key] = entries;
                }

                var plain = style.Equals("Regular", StringComparison.OrdinalIgnoreCase) || style.Equals("Bold", StringComparison.OrdinalIgnoreCase)
                    || style.Equals("Italic", StringComparison.OrdinalIgnoreCase) || style.Equals("Bold Italic", StringComparison.OrdinalIgnoreCase);
                entries.Add(new Entry(path, i, bold, italic, plain));
            }
        }
    }

    private static byte[] Read(FileStream file, long offset, int count)
    {
        var data = new byte[count];
        file.Seek(offset, SeekOrigin.Begin);
        var read = 0;
        while (read < count)
        {
            var n = file.Read(data, read, count - read);
            if (n == 0)
            {
                throw new IOException();
            }

            read += n;
        }

        return data;
    }

    private static int BigEndian(byte[] data, int offset) =>
        (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];

    private sealed class Entry
    {
        public Entry(string path, int index, bool bold, bool italic, bool plain)
        {
            Path = path;
            Index = index;
            Bold = bold;
            Italic = italic;
            Plain = plain;
        }

        public string Path { get; }

        public int Index { get; }

        public bool Bold { get; }

        public bool Italic { get; }

        public bool Plain { get; }
    }
}
