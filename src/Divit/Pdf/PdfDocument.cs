#nullable enable
namespace Divit.Pdf;

using Divit.Packaging;
using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

/// <summary>
/// A PDF document: pages drawn with embedded TrueType fonts, lines, shapes and images, written with
/// <see cref="Save"/>. Fonts carry only the characters used, so files stay small; any language the font covers
/// works, Turkish included.
/// </summary>
public sealed partial class PdfDocument
{
    private readonly List<PdfPage> _pages = new List<PdfPage>();
    private readonly List<PdfFont> _fonts = new List<PdfFont>();
    private readonly List<PdfImage> _images = new List<PdfImage>();
    private readonly Dictionary<byte[], PdfImage> _imagesByBytes = new Dictionary<byte[], PdfImage>(ReferenceComparer.Instance);

    /// <summary>
    /// File property: title, which readers show in their title bar.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// File property: author.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// File property: subject.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// File property: keywords.
    /// </summary>
    public string? Keywords { get; set; }

    /// <summary>
    /// Application name written into the file properties.
    /// </summary>
    public string Application { get; set; } = "Divit";

    /// <summary>
    /// File property: creation time (UTC); the time of saving when empty. Set it to get byte-identical files from
    /// identical input.
    /// </summary>
    public DateTime? Created { get; set; }

    /// <summary>
    /// The pages in order.
    /// </summary>
    public IReadOnlyList<PdfPage> Pages => _pages;

    /// <summary>
    /// Adds a page at the end.
    /// </summary>
    /// <param name="size">Paper size.</param>
    /// <param name="landscape">Turns the page sideways.</param>
    /// <returns>The page, to draw on.</returns>
    public PdfPage AddPage(PdfPageSize size = PdfPageSize.A4, bool landscape = false)
    {
        var (width, height) = size switch
        {
            PdfPageSize.A3 => (841.89, 1190.55),
            PdfPageSize.A5 => (419.53, 595.28),
            PdfPageSize.Letter => (612.0, 792.0),
            PdfPageSize.Legal => (612.0, 1008.0),
            _ => (595.28, 841.89),
        };
        return landscape ? AddPage(height, width) : AddPage(width, height);
    }

    /// <summary>
    /// Adds a page of any size.
    /// </summary>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    /// <returns>The page, to draw on.</returns>
    public PdfPage AddPage(double width, double height)
    {
        var page = new PdfPage(this, width, height);
        _pages.Add(page);
        return page;
    }

    /// <summary>
    /// Adds an installed font by its family name, for example "Trebuchet MS"; the closest style when the exact one
    /// is not installed. Fonts are found in the font folders of Windows, Linux and macOS.
    /// </summary>
    /// <param name="family">Family name.</param>
    /// <param name="bold">Bold.</param>
    /// <param name="italic">Italic.</param>
    /// <returns>The font, to draw and measure text with.</returns>
    /// <exception cref="ArgumentException">No font of the family is installed.</exception>
    /// <exception cref="NotSupportedException">The font has PostScript outlines.</exception>
    public PdfFont AddFont(string family, bool bold = false, bool italic = false)
    {
        var face = FontLocator.Find(family ?? throw new ArgumentNullException(nameof(family)), bold, italic)
            ?? throw new ArgumentException(Messages.Format(Strings.FontNotFound, family), nameof(family));
        return Add(face);
    }

    /// <summary>
    /// Adds a font from the bytes of its file (.ttf, or one font of a .ttc collection): for fonts that are not
    /// installed, or servers without fonts.
    /// </summary>
    /// <param name="font">Bytes of the font file.</param>
    /// <param name="index">Which font of a collection; 0 for a single font.</param>
    /// <returns>The font, to draw and measure text with.</returns>
    /// <exception cref="ArgumentException">The bytes are no TrueType font.</exception>
    /// <exception cref="NotSupportedException">The font has PostScript outlines.</exception>
    public PdfFont AddFont(byte[] font, int index = 0) => Add(new TrueTypeFont(font ?? throw new ArgumentNullException(nameof(font)), index));

    /// <summary>
    /// Adds an entry to the bookmarks readers show beside the pages, such as the headings of a report. An entry
    /// one level deeper than the one before it goes under it.
    /// </summary>
    /// <param name="title">The entry's text.</param>
    /// <param name="page">The page it jumps to.</param>
    /// <param name="y">The place on the page, from the top.</param>
    /// <param name="level">Depth, 0 for the outermost.</param>
    public void AddBookmark(string title, PdfPage page, double y = 0, int level = 0) =>
        _bookmarks.Add((title ?? string.Empty, page ?? throw new ArgumentNullException(nameof(page)), y, Math.Max(0, level)));

    private readonly List<(string Title, PdfPage Page, double Y, int Level)> _bookmarks = new List<(string, PdfPage, double, int)>();

    /// <summary>
    /// Writes the document; an existing target is replaced only once the new file is complete.
    /// </summary>
    /// <param name="path">Target file.</param>
    public void Save(string path) => AtomicFile.Save(path, Write);

    private PdfFont Add(TrueTypeFont face)
    {
        foreach (var font in _fonts)
        {
            if (ReferenceEquals(font.Face, face))
            {
                return font;
            }
        }

        var added = new PdfFont(face, "F" + OfficeText.Int(_fonts.Count + 1));
        _fonts.Add(added);
        return added;
    }

    /// <summary>
    /// The image of some bytes, decoded once per document.
    /// </summary>
    internal PdfImage Image(byte[] bytes)
    {
        if (!_imagesByBytes.TryGetValue(bytes ?? throw new ArgumentNullException(nameof(bytes)), out var image))
        {
            image = PdfImage.From(bytes);
            _imagesByBytes[bytes] = image;
            _images.Add(image);
        }

        return image;
    }

    /// <summary>
    /// The number in an image's resource name, <c>Im1</c> for the first.
    /// </summary>
    internal int ImageNumber(PdfImage image) => _images.IndexOf(image) + 1;

    /// <summary>
    /// Compares byte arrays by identity: the same array drawn twice is one image.
    /// </summary>
    private sealed class ReferenceComparer : IEqualityComparer<byte[]>
    {
        public static readonly ReferenceComparer Instance = new ReferenceComparer();

        public bool Equals(byte[]? x, byte[]? y) => ReferenceEquals(x, y);

        public int GetHashCode(byte[] obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
