#nullable enable
namespace Divit.Pdf;

using System.Collections.Generic;

/// <summary>
/// A font of a document (<see cref="PdfDocument.AddFont(string, bool, bool)"/>): an installed or given TrueType
/// font, embedded with only the characters the document uses. It measures text for laying it out by hand.
/// </summary>
public sealed class PdfFont
{
    internal PdfFont(TrueTypeFont face, string resource)
    {
        Face = face;
        Resource = resource;
    }

    /// <summary>
    /// Family name, for example "Trebuchet MS".
    /// </summary>
    public string Family => Face.Family;

    /// <summary>
    /// Style name, for example "Bold".
    /// </summary>
    public string Style => Face.Style;

    /// <summary>
    /// Height above the baseline in points at a size: where the first line's baseline goes below a top edge.
    /// </summary>
    /// <param name="size">Font size in points.</param>
    /// <returns>The height.</returns>
    public double Ascent(double size) => Face.Ascent * size / Face.UnitsPerEm;

    /// <summary>
    /// Depth below the baseline in points at a size, as a positive number.
    /// </summary>
    /// <param name="size">Font size in points.</param>
    /// <returns>The depth.</returns>
    public double Descent(double size) => -Face.Descent * size / Face.UnitsPerEm;

    /// <summary>
    /// Distance between baselines of single-spaced lines at a size, as the font recommends.
    /// </summary>
    /// <param name="size">Font size in points.</param>
    /// <returns>The distance.</returns>
    public double LineHeight(double size) => (Face.Ascent - Face.Descent + Face.LineGap) * size / Face.UnitsPerEm;

    /// <summary>
    /// Height of capital letters in points at a size, for centring text in a box.
    /// </summary>
    /// <param name="size">Font size in points.</param>
    /// <returns>The height.</returns>
    public double CapHeight(double size) => Face.CapHeight * size / Face.UnitsPerEm;

    /// <summary>
    /// Width of a text in points at a size.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <param name="size">Font size in points.</param>
    /// <returns>The width.</returns>
    public double Measure(string text, double size)
    {
        long units = 0;
        var advances = Face.Advances;
        for (var i = 0; i < text.Length; i++)
        {
            int c = text[i];
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                c = char.ConvertToUtf32(text[i], text[i + 1]);
                i++;
            }

            units += advances[Face.Glyph(c)];
        }

        return units * size / Face.UnitsPerEm;
    }

    /// <summary>
    /// Whether the font has a glyph for a character.
    /// </summary>
    /// <param name="c">The character.</param>
    /// <returns><c>true</c> when the font can draw it.</returns>
    public bool Has(char c) => Face.Glyph(c) != 0;

    internal TrueTypeFont Face { get; }

    /// <summary>
    /// The font's name in the pages' resources, such as <c>F1</c>.
    /// </summary>
    internal string Resource { get; }

    /// <summary>
    /// The glyphs used so far, with the character each stands for (for copying text out of the PDF).
    /// </summary>
    internal SortedDictionary<ushort, int> Used { get; } = new SortedDictionary<ushort, int>();
}
