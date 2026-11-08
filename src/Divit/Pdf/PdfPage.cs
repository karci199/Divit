#nullable enable
namespace Divit.Pdf;

using Divit.Packaging;
using System;
using System.Collections.Generic;

/// <summary>
/// Paper sizes.
/// </summary>
public enum PdfPageSize
{
    /// <summary>
    /// A4, 210 × 297 mm.
    /// </summary>
    A4,

    /// <summary>
    /// A3, 297 × 420 mm.
    /// </summary>
    A3,

    /// <summary>
    /// A5, 148 × 210 mm.
    /// </summary>
    A5,

    /// <summary>
    /// US Letter, 8.5 × 11 in.
    /// </summary>
    Letter,

    /// <summary>
    /// US Legal, 8.5 × 14 in.
    /// </summary>
    Legal,
}

/// <summary>
/// Where text lines up with the x it is drawn at.
/// </summary>
public enum PdfAlign
{
    /// <summary>
    /// Text starts at x.
    /// </summary>
    Left,

    /// <summary>
    /// Text is centred on x.
    /// </summary>
    Center,

    /// <summary>
    /// Text ends at x.
    /// </summary>
    Right,
}

/// <summary>
/// The pattern of a line.
/// </summary>
public enum PdfDash
{
    /// <summary>
    /// A solid line.
    /// </summary>
    Solid,

    /// <summary>
    /// Dashes.
    /// </summary>
    Dashed,

    /// <summary>
    /// Dots.
    /// </summary>
    Dotted,
}

/// <summary>
/// A page to draw on: text in the document's fonts, lines, shapes and images. Positions and sizes are in points
/// (1/72 inch, 2.835 per millimetre) from the top left corner, growing to the right and down; colours are
/// <c>#RRGGBB</c>.
/// </summary>
public sealed class PdfPage
{
    private const double Kappa = 0.5522847498;

    private readonly PdfBuffer _content = new PdfBuffer();
    private string? _fill;
    private string? _stroke;

    /// <summary>
    /// The colours as last given, so a colour drawn again is not parsed again.
    /// </summary>
    private string? _fillGiven;
    private string? _strokeGiven;
    private double _lineWidth = 1;
    private PdfDash _dash;
    private bool _round;

    internal PdfPage(PdfDocument document, double width, double height)
    {
        Document = document;
        Width = width;
        Height = height;
    }

    /// <summary>
    /// Width of the page in points.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// Height of the page in points.
    /// </summary>
    public double Height { get; }

    internal PdfDocument Document { get; }

    internal PdfBuffer Content => _content;

    /// <summary>
    /// The fonts and images the page uses, for its resources.
    /// </summary>
    internal SortedSet<string> Fonts { get; } = new SortedSet<string>(StringComparer.Ordinal);

    internal List<PdfImage> Images { get; } = new List<PdfImage>();

    /// <summary>
    /// Draws a line of text with its baseline at <paramref name="y"/>.
    /// </summary>
    /// <param name="text">The text; line feeds are not broken here.</param>
    /// <param name="x">Where it starts, is centred or ends, by <paramref name="align"/>.</param>
    /// <param name="y">The baseline.</param>
    /// <param name="font">A font of this document.</param>
    /// <param name="size">Size in points.</param>
    /// <param name="color">Colour.</param>
    /// <param name="align">How it lines up with <paramref name="x"/>.</param>
    /// <returns>This page.</returns>
#if NET
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
    public PdfPage DrawText(string text, double x, double y, PdfFont font, double size, string color = "#000000", PdfAlign align = PdfAlign.Left)
    {
        if (font == null)
        {
            throw new ArgumentNullException(nameof(font));
        }

        if (string.IsNullOrEmpty(text))
        {
            return this;
        }

        var left = align == PdfAlign.Left ? x : x - (font.Measure(text, size) * (align == PdfAlign.Center ? 0.5 : 1));
        Fonts.Add(font.Resource);
        Fill(color);
        var b = _content.Append("BT /").Append(font.Resource).Append(' ').Number(size).Append("Tf ").Number(left).Number(Height - y).Append("Td ");
        Show(b, text, font);
        b.Append(" ET\n");
        return this;
    }

    /// <summary>
    /// Fills a rectangle.
    /// </summary>
    /// <param name="x">Left edge.</param>
    /// <param name="y">Top edge.</param>
    /// <param name="width">Width.</param>
    /// <param name="height">Height.</param>
    /// <param name="color">Fill colour.</param>
    /// <returns>This page.</returns>
    public PdfPage FillRectangle(double x, double y, double width, double height, string color)
    {
        Fill(color);
        Rectangle(x, y, width, height).Append("f\n");
        return this;
    }

    /// <summary>
    /// Draws the outline of a rectangle.
    /// </summary>
    /// <param name="x">Left edge.</param>
    /// <param name="y">Top edge.</param>
    /// <param name="width">Width.</param>
    /// <param name="height">Height.</param>
    /// <param name="color">Line colour.</param>
    /// <param name="lineWidth">Line width in points.</param>
    /// <returns>This page.</returns>
    public PdfPage StrokeRectangle(double x, double y, double width, double height, string color, double lineWidth = 0.5)
    {
        Stroke(color, lineWidth, PdfDash.Solid, false);
        Rectangle(x, y, width, height).Append("S\n");
        return this;
    }

    /// <summary>
    /// Draws a straight line.
    /// </summary>
    /// <param name="x1">Start, from the left.</param>
    /// <param name="y1">Start, from the top.</param>
    /// <param name="x2">End, from the left.</param>
    /// <param name="y2">End, from the top.</param>
    /// <param name="color">Colour.</param>
    /// <param name="lineWidth">Width in points.</param>
    /// <param name="dash">Pattern.</param>
    /// <returns>This page.</returns>
    public PdfPage DrawLine(double x1, double y1, double x2, double y2, string color, double lineWidth = 0.5, PdfDash dash = PdfDash.Solid)
    {
        Stroke(color, lineWidth, dash, false);
        _content.Number(x1).Number(Height - y1).Append("m ").Number(x2).Number(Height - y2).Append("l S\n");
        return this;
    }

    /// <summary>
    /// Draws connected straight lines through points, with round joins: a measured curve.
    /// </summary>
    /// <param name="points">The points in order.</param>
    /// <param name="color">Colour.</param>
    /// <param name="lineWidth">Width in points.</param>
    /// <param name="dash">Pattern.</param>
    /// <returns>This page.</returns>
    public PdfPage DrawPolyline(IReadOnlyList<(double X, double Y)> points, string color, double lineWidth = 1, PdfDash dash = PdfDash.Solid)
    {
        if (points.Count < 2)
        {
            return this;
        }

        Stroke(color, lineWidth, dash, true);
        Path(points).Append("S\n");
        return this;
    }

    /// <summary>
    /// Fills the shape the points enclose.
    /// </summary>
    /// <param name="points">The corners in order.</param>
    /// <param name="color">Fill colour.</param>
    /// <returns>This page.</returns>
    public PdfPage FillPolygon(IReadOnlyList<(double X, double Y)> points, string color)
    {
        if (points.Count < 3)
        {
            return this;
        }

        Fill(color);
        Path(points).Append("h f\n");
        return this;
    }

    /// <summary>
    /// Fills an ellipse, or a circle with equal radii.
    /// </summary>
    /// <param name="cx">Centre, from the left.</param>
    /// <param name="cy">Centre, from the top.</param>
    /// <param name="rx">Horizontal radius.</param>
    /// <param name="ry">Vertical radius.</param>
    /// <param name="color">Fill colour.</param>
    /// <returns>This page.</returns>
    public PdfPage FillEllipse(double cx, double cy, double rx, double ry, string color)
    {
        Fill(color);
        Ellipse(cx, cy, rx, ry).Append("f\n");
        return this;
    }

    /// <summary>
    /// Draws the outline of an ellipse, or of a circle with equal radii.
    /// </summary>
    /// <param name="cx">Centre, from the left.</param>
    /// <param name="cy">Centre, from the top.</param>
    /// <param name="rx">Horizontal radius.</param>
    /// <param name="ry">Vertical radius.</param>
    /// <param name="color">Line colour.</param>
    /// <param name="lineWidth">Line width in points.</param>
    /// <returns>This page.</returns>
    public PdfPage StrokeEllipse(double cx, double cy, double rx, double ry, string color, double lineWidth = 0.5)
    {
        Stroke(color, lineWidth, PdfDash.Solid, false);
        Ellipse(cx, cy, rx, ry).Append("S\n");
        return this;
    }

    /// <summary>
    /// Draws an image (PNG or JPEG) into a box; the same bytes drawn again are stored once. With a height of 0 it
    /// keeps the image's proportions.
    /// </summary>
    /// <param name="image">Bytes of the image file.</param>
    /// <param name="x">Left edge.</param>
    /// <param name="y">Top edge.</param>
    /// <param name="width">Width.</param>
    /// <param name="height">Height; 0 to follow the width.</param>
    /// <returns>This page.</returns>
    /// <exception cref="ArgumentException">The bytes are no image.</exception>
    /// <exception cref="NotSupportedException">The image is a GIF or an interlaced PNG.</exception>
    public PdfPage DrawImage(byte[] image, double x, double y, double width, double height = 0)
    {
        var picture = Document.Image(image);
        if (!Images.Contains(picture))
        {
            Images.Add(picture);
        }

        if (height <= 0)
        {
            height = width * picture.Height / picture.Width;
        }

        _content.Append("q ").Number(width).Append("0 0 ").Number(height).Number(x).Number(Height - y - height).Append("cm /Im")
            .Integer(Document.ImageNumber(picture)).Append(" Do Q\n");
        return this;
    }

    /// <summary>
    /// Makes an area of the page a link to a web address: a click opens it.
    /// </summary>
    /// <param name="x">Left edge.</param>
    /// <param name="y">Top edge.</param>
    /// <param name="width">Width.</param>
    /// <param name="height">Height.</param>
    /// <param name="url">The address, for example <c>https://example.com</c> or <c>mailto:</c>.</param>
    /// <returns>This page.</returns>
    public PdfPage AddLink(double x, double y, double width, double height, string url)
    {
        Links.Add(new PdfLink(x, y, width, height, url ?? throw new ArgumentNullException(nameof(url)), null, 0));
        return this;
    }

    /// <summary>
    /// Makes an area of the page a link to a place in the document: a click jumps there.
    /// </summary>
    /// <param name="x">Left edge.</param>
    /// <param name="y">Top edge.</param>
    /// <param name="width">Width.</param>
    /// <param name="height">Height.</param>
    /// <param name="target">The page to jump to.</param>
    /// <param name="targetY">The place on it, from the top.</param>
    /// <returns>This page.</returns>
    public PdfPage AddLink(double x, double y, double width, double height, PdfPage target, double targetY = 0)
    {
        Links.Add(new PdfLink(x, y, width, height, null, target ?? throw new ArgumentNullException(nameof(target)), targetY));
        return this;
    }

    /// <summary>
    /// Draws text turned about its start by an angle in degrees, counter-clockwise: a watermark.
    /// </summary>
    internal void DrawTurnedText(string text, double x, double y, PdfFont font, double size, string color, double degrees)
    {
        Fonts.Add(font.Resource);
        Fill(color);
        var radians = degrees * Math.PI / 180;
        double cos = Math.Cos(radians), sin = Math.Sin(radians);
        var b = _content.Append("BT /").Append(font.Resource).Append(' ').Number(size).Append("Tf ").Number(cos).Number(sin).Number(-sin).Number(cos)
            .Number(x).Number(Height - y).Append("Tm ");
        Show(b, text, font);
        b.Append(" ET\n");
    }

    /// <summary>
    /// The link areas of the page.
    /// </summary>
    internal List<PdfLink> Links { get; } = new List<PdfLink>();

    /// <summary>
    /// Shows a text: its glyph ids as hex, recorded as used; with kerning an array that moves the glyphs of kerned
    /// pairs together (<c>TJ</c>), else one string (<c>Tj</c>).
    /// </summary>
#if NET
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
    private static void Show(PdfBuffer b, string text, PdfFont font)
    {
        var face = font.Face;
        var used = font.Used;
        var kerns = face.Kerning.Any;
        var scale = 1000.0 / face.UnitsPerEm;
        b.Append(kerns ? "[<" : "<");
        ushort previous = 0;
        for (var i = 0; i < text.Length; i++)
        {
            int c = text[i];
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
            {
                c = char.ConvertToUtf32(text[i], text[i + 1]);
                i++;
            }

            var glyph = face.Glyph(c);
            if (glyph != 0 && !used.ContainsKey(glyph))
            {
                used[glyph] = c;
            }

            if (kerns && previous != 0)
            {
                var kern = font.Kern(previous, glyph);
                if (kern != 0)
                {
                    // A positive number moves the next glyph left: the opposite sign of the adjustment.
                    b.Append("> ").Number(-kern * scale).Append('<');
                }
            }

            b.Hex4(glyph);
            previous = glyph;
        }

        b.Append(kerns ? ">] TJ" : "> Tj");
    }

    private PdfBuffer Rectangle(double x, double y, double width, double height) =>
        _content.Number(x).Number(Height - y - height).Number(width).Number(height).Append("re ");

    private PdfBuffer Path(IReadOnlyList<(double X, double Y)> points)
    {
        _content.Number(points[0].X).Number(Height - points[0].Y).Append("m ");
        for (var i = 1; i < points.Count; i++)
        {
            _content.Number(points[i].X).Number(Height - points[i].Y).Append("l ");
        }

        return _content;
    }

    /// <summary>
    /// An ellipse as four Bézier curves.
    /// </summary>
    private PdfBuffer Ellipse(double cx, double cy, double rx, double ry)
    {
        var y = Height - cy;
        double ox = rx * Kappa, oy = ry * Kappa;
        var b = _content.Number(cx + rx).Number(y).Append("m ");
        b.Number(cx + rx).Number(y + oy).Number(cx + ox).Number(y + ry).Number(cx).Number(y + ry).Append("c ");
        b.Number(cx - ox).Number(y + ry).Number(cx - rx).Number(y + oy).Number(cx - rx).Number(y).Append("c ");
        b.Number(cx - rx).Number(y - oy).Number(cx - ox).Number(y - ry).Number(cx).Number(y - ry).Append("c ");
        b.Number(cx + ox).Number(y - ry).Number(cx + rx).Number(y - oy).Number(cx + rx).Number(y).Append("c ");
        return b;
    }

    /// <summary>
    /// Sets the fill colour, only when it changes.
    /// </summary>
    private void Fill(string color)
    {
        if (color == _fillGiven)
        {
            return;
        }

        _fillGiven = color;
        var rgb = OfficeText.Rgb(color);
        if (rgb == _fill)
        {
            return;
        }

        _fill = rgb;
        Color(rgb).Append("rg\n");
    }

    /// <summary>
    /// Sets the line colour, width, pattern and joins, only what changes.
    /// </summary>
    private void Stroke(string color, double width, PdfDash dash, bool round)
    {
        if (color != _strokeGiven)
        {
            _strokeGiven = color;
            var rgb = OfficeText.Rgb(color);
            if (rgb != _stroke)
            {
                _stroke = rgb;
                Color(rgb).Append("RG\n");
            }
        }

        var widthChanged = width != _lineWidth;
        if (widthChanged)
        {
            _lineWidth = width;
            _content.Number(width).Append("w\n");
        }

        // Dashes scale with the line, so a new width needs the pattern again.
        if (dash != _dash || (dash != PdfDash.Solid && widthChanged))
        {
            _dash = dash;
            _content.Append(dash == PdfDash.Solid ? "[] 0 d\n" : "[");
            if (dash == PdfDash.Dashed)
            {
                _content.Number(width * 4).Number(width * 3).Append("] 0 d\n");
            }
            else if (dash == PdfDash.Dotted)
            {
                _content.Number(width).Number(width * 2).Append("] 0 d\n");
            }
        }

        if (round != _round)
        {
            _round = round;
            _content.Append(round ? "1 j 1 J\n" : "0 j 0 J\n");
        }
    }

    private PdfBuffer Color(string rgb)
    {
        for (var i = 0; i < 6; i += 2)
        {
            _content.Number(Convert.ToInt32(rgb.Substring(i, 2), 16) / 255.0);
        }

        return _content;
    }
}

/// <summary>
/// A link area of a page: to a web address, or to a place on a page of the document.
/// </summary>
internal sealed class PdfLink
{
    public PdfLink(double x, double y, double width, double height, string? url, PdfPage? target, double targetY)
    {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Url = url;
        Target = target;
        TargetY = targetY;
    }

    public double X { get; }

    public double Y { get; }

    public double Width { get; }

    public double Height { get; }

    public string? Url { get; }

    public PdfPage? Target { get; }

    public double TargetY { get; }
}
