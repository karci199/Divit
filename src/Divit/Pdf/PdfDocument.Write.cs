#nullable enable
namespace Divit.Pdf;

using Divit.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

/// <content>
/// The file: catalog, pages, images and fonts as numbered objects, then the cross-reference table.
/// </content>
public sealed partial class PdfDocument
{
    /// <summary>
    /// Writes the objects in the order of their ids: catalog, page tree, information, then per page its object and
    /// contents, the images with their masks, and the fonts.
    /// </summary>
    private void Write(Stream stream)
    {
        var file = new PdfFile(stream);
        int catalog = file.Reserve(), tree = file.Reserve(), info = file.Reserve();
        var pageIds = new int[_pages.Count];
        var contentIds = new int[_pages.Count];
        for (var i = 0; i < _pages.Count; i++)
        {
            pageIds[i] = file.Reserve();
            contentIds[i] = file.Reserve();
        }

        var imageIds = new Dictionary<PdfImage, (int Image, int Mask)>();
        foreach (var image in _images)
        {
            imageIds[image] = (file.Reserve(), image.Mask != null ? file.Reserve() : 0);
        }

        var fontIds = new Dictionary<string, int>(StringComparer.Ordinal);
        var fontObjects = new List<(PdfFont Font, int Type0, int Cid, int Descriptor, int Program, int Unicode)>();
        foreach (var font in _fonts)
        {
            if (font.Used.Count == 0 && !Used(font))
            {
                continue;
            }

            var ids = (font, file.Reserve(), file.Reserve(), file.Reserve(), file.Reserve(), file.Reserve());
            fontObjects.Add(ids);
            fontIds[font.Resource] = ids.Item2;
        }

        file.Begin(catalog).Append("<< /Type /Catalog /Pages ").Integer(tree).Append(" 0 R >>");
        file.End();
        var kids = file.Begin(tree).Append("<< /Type /Pages /Kids [");
        foreach (var id in pageIds)
        {
            kids.Integer(id).Append(" 0 R ");
        }

        kids.Append("] /Count ").Integer(_pages.Count).Append(" >>");
        file.End();
        WriteInfo(file.Begin(info));
        file.End();

        for (var i = 0; i < _pages.Count; i++)
        {
            var page = _pages[i];
            var b = file.Begin(pageIds[i]).Append("<< /Type /Page /Parent ").Integer(tree).Append(" 0 R /MediaBox [0 0 ").Number(page.Width).Number(page.Height)
                .Append("] /Resources << ");
            if (page.Fonts.Count > 0)
            {
                b.Append("/Font << ");
                foreach (var name in page.Fonts)
                {
                    b.Append('/').Append(name).Append(' ').Integer(fontIds[name]).Append(" 0 R ");
                }

                b.Append(">> ");
            }

            if (page.Images.Count > 0)
            {
                b.Append("/XObject << ");
                foreach (var image in page.Images)
                {
                    b.Append("/Im").Integer(ImageNumber(image)).Append(' ').Integer(imageIds[image].Image).Append(" 0 R ");
                }

                b.Append(">> ");
            }

            b.Append(">> /Contents ").Integer(contentIds[i]).Append(" 0 R >>");
            file.End();
            file.Stream(contentIds[i], string.Empty, page.Content.Data, page.Content.Length);
        }

        foreach (var image in _images)
        {
            var (id, mask) = imageIds[image];
            var entries = new StringBuilder("/Type /XObject /Subtype /Image /Width ").Append(OfficeText.Int(image.Width)).Append(" /Height ")
                .Append(OfficeText.Int(image.Height)).Append(' ').Append(image.Entries);
            if (mask != 0)
            {
                entries.Append("/SMask ").Append(OfficeText.Int(mask)).Append(" 0 R ");
            }

            file.Stream(id, entries.ToString(), image.Data, image.Data.Length, image.Compress);
            if (mask != 0)
            {
                var m = image.Mask!;
                file.Stream(mask, "/Type /XObject /Subtype /Image /Width " + OfficeText.Int(m.Width) + " /Height " + OfficeText.Int(m.Height) + " " + m.Entries, m.Data, m.Data.Length, m.Compress);
            }
        }

        foreach (var (font, type0, cid, descriptor, program, unicode) in fontObjects)
        {
            WriteFont(file, font, type0, cid, descriptor, program, unicode);
        }

        file.Close(catalog, info);
    }

    /// <summary>
    /// Whether a font with no glyphs drawn still appears on a page (text of characters it lacks).
    /// </summary>
    private bool Used(PdfFont font)
    {
        foreach (var page in _pages)
        {
            if (page.Fonts.Contains(font.Resource))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// The information dictionary: the file properties and the times.
    /// </summary>
    private void WriteInfo(PdfBuffer b)
    {
        var time = (Created ?? DateTime.UtcNow).ToUniversalTime().ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
        b.Append("<< ");
        TextEntry(b, "Title", Title);
        TextEntry(b, "Author", Author);
        TextEntry(b, "Subject", Subject);
        TextEntry(b, "Keywords", Keywords);
        TextEntry(b, "Creator", Application);
        TextEntry(b, "Producer", "Divit");
        b.Append("/CreationDate (D:").Append(time).Append("Z) /ModDate (D:").Append(time).Append("Z) >>");
    }

    /// <summary>
    /// A text entry as UTF-16 with a byte order mark, which every reader shows in any language.
    /// </summary>
    private static void TextEntry(PdfBuffer b, string key, string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        b.Append('/').Append(key).Append(" <FEFF");
        foreach (var c in value!)
        {
            b.Hex4(c);
        }

        b.Append("> ");
    }

    /// <summary>
    /// A font as Word and Acrobat embed TrueType for any language: a Type 0 font over a CID font whose ids are the
    /// glyph ids, with the subset font program, the widths of the used glyphs, and a map back to the characters.
    /// </summary>
    private static void WriteFont(PdfFile file, PdfFont font, int type0, int cid, int descriptor, int program, int unicode)
    {
        var face = font.Face;
        var glyphs = new SortedSet<ushort>(font.Used.Keys);
        var subset = FontSubset.Build(face, glyphs, font.Used);
        var name = Tag(font.Used) + "+" + face.PostScriptName;
        double scale = 1000.0 / face.UnitsPerEm;

        file.Begin(type0).Append("<< /Type /Font /Subtype /Type0 /BaseFont /").Append(name).Append(" /Encoding /Identity-H /DescendantFonts [")
            .Integer(cid).Append(" 0 R] /ToUnicode ").Integer(unicode).Append(" 0 R >>");
        file.End();

        // Widths of the used glyphs, in runs of consecutive ids.
        var b = file.Begin(cid).Append("<< /Type /Font /Subtype /CIDFontType2 /BaseFont /").Append(name)
            .Append(" /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ").Integer(descriptor)
            .Append(" 0 R /CIDToGIDMap /Identity /DW ").Integer((long)Math.Round(face.Advances[0] * scale)).Append(" /W [");
        var previous = -2;
        foreach (var glyph in font.Used.Keys)
        {
            if (glyph != previous + 1)
            {
                if (previous >= 0)
                {
                    b.Append("] ");
                }

                b.Integer(glyph).Append(" [");
            }

            b.Integer((long)Math.Round(face.Advances[glyph] * scale)).Append(' ');
            previous = glyph;
        }

        b.Append(previous >= 0 ? "]] >>" : "] >>");
        file.End();

        var flags = 32 | (face.FixedPitch ? 1 : 0) | (face.Italic ? 64 : 0);
        file.Begin(descriptor).Append("<< /Type /FontDescriptor /FontName /").Append(name).Append(" /Flags ").Integer(flags)
            .Append(" /FontBBox [").Number(face.XMin * scale).Number(face.YMin * scale).Number(face.XMax * scale).Number(face.YMax * scale)
            .Append("] /ItalicAngle ").Number(face.ItalicAngle).Append("/Ascent ").Number(face.Ascent * scale).Append("/Descent ").Number(face.Descent * scale)
            .Append("/CapHeight ").Number(face.CapHeight * scale).Append("/StemV ").Integer(face.Bold || face.Weight >= 600 ? 120 : 80)
            .Append(" /FontFile2 ").Integer(program).Append(" 0 R >>");
        file.End();

        file.Stream(program, "/Length1 " + OfficeText.Int(subset.Length) + " ", subset, subset.Length);
        var cmap = ToUnicode(font.Used);
        file.Stream(unicode, string.Empty, cmap.Data, cmap.Length);
    }

    /// <summary>
    /// The map from glyph ids back to characters, so text can be searched and copied out of the PDF.
    /// </summary>
    private static PdfBuffer ToUnicode(SortedDictionary<ushort, int> used)
    {
        var b = new PdfBuffer();
        b.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n")
            .Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");
        var entries = new List<KeyValuePair<ushort, int>>(used);
        for (var start = 0; start < entries.Count; start += 100)
        {
            var count = Math.Min(100, entries.Count - start);
            b.Integer(count).Append(" beginbfchar\n");
            for (var i = start; i < start + count; i++)
            {
                b.Append('<');
                b.Hex4(entries[i].Key);
                b.Append("> <");
                var text = char.ConvertFromUtf32(entries[i].Value);
                foreach (var c in text)
                {
                    b.Hex4(c);
                }

                b.Append(">\n");
            }

            b.Append("endbfchar\n");
        }

        b.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        return b;
    }

    /// <summary>
    /// The six capital letters that mark a subset font, derived from the glyphs it holds.
    /// </summary>
    private static string Tag(SortedDictionary<ushort, int> used)
    {
        uint hash = 2166136261;
        foreach (var glyph in used.Keys)
        {
            hash = (hash ^ glyph) * 16777619;
        }

        var tag = new char[6];
        for (var i = 0; i < 6; i++)
        {
            tag[i] = (char)('A' + (hash % 26));
            hash /= 26;
        }

        return new string(tag);
    }
}
