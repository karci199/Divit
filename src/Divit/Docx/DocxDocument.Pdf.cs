#nullable enable
namespace Divit.Docx;

using Divit.Pdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

/// <content>
/// The document laid out as PDF: the same sections, styles, headings, lists, tables, images, headers and footers,
/// broken into lines and pages by Divit instead of Word.
/// </content>
public sealed partial class DocxDocument
{
    /// <summary>
    /// Lays the document out as PDF pages: what Word would show, with the page numbers of the contents filled in.
    /// More can be drawn on the pages before saving.
    /// </summary>
    /// <remarks>
    /// <para>Fonts come from the installed fonts by the look's names (Arial when one is missing) and are embedded as
    /// subsets. Lines break at spaces, pages between lines and table rows; header rows repeat on every page a table
    /// spans, and headings stay with the paragraph after them.</para>
    /// <para>Equations are written as text in a line, tracked changes as accepted, check boxes as boxes. Streamed
    /// sections are not laid out: their content exists only while a Word file is saved.</para>
    /// </remarks>
    /// <returns>The PDF document.</returns>
    public PdfDocument ToPdf() => new PdfLayout(this).Render();

    /// <summary>
    /// Saves the document as PDF (<see cref="ToPdf"/>); an existing target is replaced only once the new file is
    /// complete.
    /// </summary>
    /// <param name="path">Target file.</param>
    public void SavePdf(string path) => ToPdf().Save(path);

    /// <summary>
    /// Points per centimetre.
    /// </summary>
    private const double Cm = 72 / 2.54;

    /// <summary>
    /// The layout of one document: paragraphs become lines, lines and rows become bands, bands fill pages.
    /// </summary>
    private sealed class PdfLayout
    {
        private static readonly string[] Fallbacks = { "Segoe UI Symbol", "Segoe UI", "Arial Unicode MS", "DejaVu Sans", "Arial" };

        private readonly DocxDocument _doc;
        private readonly PdfDocument _pdf;
        private readonly Dictionary<(string Family, bool Bold, bool Italic), PdfFont> _fonts = new Dictionary<(string, bool, bool), PdfFont>();
        private readonly Dictionary<DocxParagraph, string> _numbers = new Dictionary<DocxParagraph, string>();
        private readonly List<(DocxParagraph Paragraph, int Level, string Number)> _headings = new List<(DocxParagraph, int, string)>();

        /// <summary>
        /// Where each paragraph that starts a band landed in the last pass: page index and top.
        /// </summary>
        private Dictionary<DocxParagraph, (int Page, double Y)> _placed = new Dictionary<DocxParagraph, (int, double)>();

        private Dictionary<DocxList, int[]> _counters = new Dictionary<DocxList, int[]>();
        private int _footnotes;
        private DocxParagraph? _previous;
        private string? _next;

        public PdfLayout(DocxDocument doc)
        {
            _doc = doc;
            _pdf = new PdfDocument { Title = doc.Title, Author = doc.Author, Subject = doc.Subject, Keywords = doc.Keywords, Application = doc.Application, Created = doc.Created };
            doc.Headings(_numbers);
            foreach (var section in doc._sections)
            {
                foreach (var block in section.Blocks)
                {
                    if (block is DocxParagraph p && HeadingLevel(p) is int level)
                    {
                        _headings.Add((p, level, _numbers.TryGetValue(p, out var number) ? number : string.Empty));
                    }
                }
            }
        }

        private DocxLook Look => _doc.Look;

        public PdfDocument Render()
        {
            // Laid out twice when the document lists pages (contents, page references): the second pass knows them.
            var pages = Flow();
            if (NeedsPages())
            {
                _placed = Placements(pages);
                pages = Flow();
            }

            Emit(pages);
            return _pdf;
        }

        private bool NeedsPages()
        {
            foreach (var section in _doc._sections)
            {
                foreach (var paragraph in Paragraphs(section))
                {
                    foreach (var run in paragraph.Runs)
                    {
                        if (run.ReferenceTarget != null && run.ReferenceKind == DocxReferenceKind.Page) { return true; }
                    }
                }

                foreach (var block in section.Blocks)
                {
                    if (block is DocxTableOfContents) { return true; }
                }
            }

            return false;
        }

        private static int? HeadingLevel(DocxParagraph p) =>
            p.Style.Length == 8 && p.Style.StartsWith("Heading", StringComparison.Ordinal) && p.Style[7] >= '1' && p.Style[7] <= '6' ? p.Style[7] - '0' : (int?)null;

        // ---------------------------------------------------------------------------------------------- fonts

        private PdfFont Font(string? family, bool bold, bool italic)
        {
            var name = family ?? Look.BodyFont;
            if (_fonts.TryGetValue((name, bold, italic), out var font))
            {
                return font;
            }

            foreach (var candidate in new[] { name, Look.BodyFont, "Arial", "Liberation Sans", "DejaVu Sans" })
            {
                try
                {
                    font = _pdf.AddFont(candidate, bold, italic);
                    break;
                }
                catch (ArgumentException)
                {
                }
                catch (NotSupportedException)
                {
                }
            }

            _fonts[(name, bold, italic)] = font ?? throw new InvalidOperationException(Divit.Properties.Messages.Format(Divit.Properties.Strings.FontNotFound, name));
            return font;
        }

        /// <summary>
        /// A font that can draw every character of a word: the word's own, else a fallback that has them.
        /// </summary>
        private PdfFont Covering(PdfFont font, string text, bool bold, bool italic)
        {
            if (Covers(font, text))
            {
                return font;
            }

            foreach (var family in Fallbacks)
            {
                PdfFont candidate;
                try
                {
                    candidate = Font(family, bold, italic);
                }
                catch (InvalidOperationException)
                {
                    continue;
                }

                if (Covers(candidate, text))
                {
                    return candidate;
                }
            }

            return font;
        }

        private static bool Covers(PdfFont font, string text)
        {
            foreach (var c in text)
            {
                if (!char.IsSurrogate(c) && !char.IsWhiteSpace(c) && !font.Has(c))
                {
                    return false;
                }
            }

            return true;
        }

        // ---------------------------------------------------------------------------------------------- styles

        /// <summary>
        /// The look of a paragraph: its style from the house look, then its own settings.
        /// </summary>
        private Style Resolve(DocxParagraph p, DocxAlignment? fallback, Cell? cell)
        {
            var look = Look;
            var s = new Style { Family = look.BodyFont, Size = look.BodySize, Color = look.BodyColor, After = look.SpaceAfter, Line = look.LineSpacing };
            var heading = look.HeadingFont ?? look.BodyFont;
            switch (p.Style)
            {
                case DocxStyle.Title:
                    s.Family = heading; s.Size = look.TitleSize; s.Color = look.HeadingColor; s.Bold = true; s.After = 12; s.Line = 1;
                    break;
                case DocxStyle.Subtitle:
                    s.Family = heading; s.Size = look.BodySize + 3; s.Color = look.HeadingColor; s.After = 12;
                    break;
                case DocxStyle.Caption:
                    s.Italic = true; s.Color = look.CaptionColor; s.Size = Math.Max(8, look.BodySize - 1); s.Before = 3; s.After = 9; s.Line = 1;
                    break;
                case DocxStyle.NoSpacing:
                    s.After = 0; s.Line = 1;
                    break;
                case TocHeadingStyle:
                    s.Family = heading; s.Size = look.HeadingSize(1); s.Color = look.HeadingColor; s.Bold = true; s.Before = 18; s.After = 6; s.KeepNext = true;
                    break;
                case FiguresStyle:
                    s.After = 3;
                    break;
                case DocxStyle.ListParagraph:
                    s.Left = 36; s.Contextual = true;
                    break;
                case FootnoteTextStyle:
                    s.Size = Math.Max(7, look.BodySize - 2); s.After = 0; s.Line = 1;
                    break;
                case BibliographyStyleId:
                    s.Left = 28.35; s.First = -28.35;
                    break;
                default:
                    if (HeadingLevel(p) is int level)
                    {
                        s.Family = heading; s.Size = look.HeadingSize(level); s.Color = look.HeadingColor; s.Bold = level <= 3;
                        s.Before = level == 1 ? 18 : 12; s.After = 6; s.KeepNext = true; s.KeepLines = true;
                    }
                    else if (p.Style.StartsWith("TOC", StringComparison.Ordinal) && p.Style.Length == 4 && char.IsDigit(p.Style[3]))
                    {
                        var tocLevel = p.Style[3] - '0';
                        s.After = tocLevel == 1 ? 6 : 3; s.Left = 11 * (tocLevel - 1); s.Bold = tocLevel == 1;
                    }

                    break;
            }

            if (cell != null)
            {
                // The table style's paragraphs: compact, single spaced; the header row bold in its colour.
                s.Before = 2; s.After = 2; s.Line = 1;
                if (cell.Header)
                {
                    s.Bold = true;
                    s.Color = cell.HeaderColor;
                }
            }

            if (p.Style == DocxStyle.NoSpacing || (cell != null && p.Style == DocxStyle.Normal))
            {
                s.Before = cell != null ? 2 : 0;
                s.After = cell != null ? 2 : 0;
            }

            s.Align = p.Alignment ?? fallback ?? DocxAlignment.Left;
            if (p.SpaceBefore is double before) { s.Before = before; }
            if (p.SpaceAfter is double after) { s.After = after; }
            if (p.LineSpacing is double line) { s.Line = line; }
            if (p.LeftIndent != 0) { s.Left = p.LeftIndent * Cm; }
            if (p.RightIndent != 0) { s.Right = p.RightIndent * Cm; }
            if (p.FirstLineIndent != 0) { s.First = p.FirstLineIndent * Cm; }
            s.KeepNext |= p.KeepWithNext;
            s.KeepLines |= p.KeepTogether;
            if (p.List != null)
            {
                s.Left = 36 * (p.ListLevel + 1);
                s.First = -18;
            }

            if (_doc.NumberHeadings && HeadingLevel(p) is int numbered)
            {
                var indent = (432 + (144 * (numbered - 1))) / 20.0;
                s.Left = indent;
                s.First = -indent;
            }

            return s;
        }

        // ---------------------------------------------------------------------------------------------- paragraphs

        /// <summary>
        /// The pieces of a paragraph with their fonts and widths: words, spaces, tabs, breaks, images, boxes.
        /// </summary>
        private List<Atom> Atoms(DocxParagraph p, Style s, Fields? fields, List<Band> notes, double width)
        {
            var atoms = new List<Atom>();
            if (p.List is DocxList list)
            {
                atoms.Add(Word(ListMarker(list, p.ListLevel), s, null, null));
                atoms.Add(new Atom { Kind = AtomKind.Tab });
            }

            if (_doc.NumberHeadings && _numbers.TryGetValue(p, out var number) && number.Length > 0)
            {
                atoms.Add(Word(number, s, null, null));
                atoms.Add(new Atom { Kind = AtomKind.Tab });
            }

            foreach (var run in p.Runs)
            {
                if (run.Revision == DocxRevision.Deletion)
                {
                    continue;
                }

                if (run.Image != null)
                {
                    atoms.Add(new Atom { Kind = AtomKind.Image, Image = run.Image, Width = run.ImageWidth / 12700.0, Ascent = run.ImageHeight / 12700.0 });
                    continue;
                }

                if (run.CheckBox is bool ticked)
                {
                    var size = run.Font?.Size > 0 ? run.Font.Value.Size : s.Size;
                    atoms.Add(new Atom { Kind = AtomKind.Box, Checked = ticked, Width = size * 0.8, Ascent = size * 0.75, Size = size, Color = run.Font?.Color ?? s.Color });
                    continue;
                }

                if (run.PageBreak)
                {
                    atoms.Add(new Atom { Kind = AtomKind.PageBreak });
                    continue;
                }

                if (run.Tab) { atoms.Add(new Atom { Kind = AtomKind.Tab }); }
                if (run.LineBreak) { atoms.Add(new Atom { Kind = AtomKind.Break }); }

                string? text = run.Text;
                DocxFont? font = run.Font;
                string? url = run.LinkUrl;
                DocxParagraph? target = run.LinkTarget;
                if (run.LinkUrl != null || run.LinkTarget != null)
                {
                    var f = font ?? default;
                    f.Color ??= Look.LinkColor;
                    f.Underline = true;
                    font = f;
                }
                else if (run.Field is DocxField field)
                {
                    text = FieldText(field, fields);
                }
                else if (run.Code != null)
                {
                    text = run.Result;
                }
                else if (run.ReferenceTarget is DocxParagraph referenced)
                {
                    text = ReferenceText(run, referenced);
                    target = referenced;
                }
                else if (run.Citation != null)
                {
                    text = _doc.CitationText(run.Citation, run.CitedPages);
                }
                else if (run.Math != null)
                {
                    var f = font ?? default;
                    f.Name = "Cambria Math";
                    f.Italic = true;
                    font = f;
                    text = run.Math.ToString();
                }
                else if (run.Footnote != null)
                {
                    if (fields == null)
                    {
                        var mark = Interlocked(++_footnotes);
                        var raised = font ?? default;
                        raised.Script = DocxScript.Superscript;
                        AddText(atoms, mark, s, raised, null, null);
                        var note = new DocxParagraph(FootnoteTextStyle);
                        note.Add(mark, new DocxFont { Script = DocxScript.Superscript }).Add(" " + run.Footnote);
                        notes.AddRange(Paragraph(note, width, null, null, null, new List<Band>()));
                    }

                    continue;
                }

                if (!string.IsNullOrEmpty(text))
                {
                    AddText(atoms, text!, s, font, url, target ?? p.ContentsTarget);
                }
            }

            return atoms;
        }

        private static string Interlocked(int value) => value.ToString(CultureInfo.InvariantCulture);

        private void AddText(List<Atom> atoms, string text, Style s, DocxFont? font, string? url, DocxParagraph? target)
        {
            if (font is DocxFont f && (f.AllCaps || f.SmallCaps))
            {
                text = text.ToUpperInvariant();
            }

            var start = 0;
            for (var i = 0; i <= text.Length; i++)
            {
                if (i == text.Length || text[i] == ' ')
                {
                    if (i > start)
                    {
                        atoms.Add(Word(text.Substring(start, i - start), s, font, url, target));
                    }

                    if (i < text.Length)
                    {
                        var space = Word(" ", s, font, url, target);
                        space.Kind = AtomKind.Space;
                        atoms.Add(space);
                    }

                    start = i + 1;
                }
            }
        }

        private Atom Word(string text, Style s, DocxFont? font, string? url, DocxParagraph? target = null)
        {
            var f = font ?? default;
            var bold = s.Bold || f.Bold;
            var italic = s.Italic || f.Italic;
            var size = f.Size > 0 ? f.Size : s.Size;
            if (f.SmallCaps) { size *= 0.8; }
            var rise = 0.0;
            if (f.Script != DocxScript.None)
            {
                rise = f.Script == DocxScript.Superscript ? size * 0.33 : -size * 0.14;
                size *= 0.65;
            }

            var face = Covering(Font(f.Name ?? s.Family, bold, italic), text, bold, italic);
            return new Atom
            {
                Kind = AtomKind.Word,
                Text = text,
                Font = face,
                Size = size,
                Color = f.Color ?? s.Color,
                Underline = f.Underline,
                Strike = f.Strikethrough,
                Back = f.Background,
                Rise = rise,
                Width = face.Measure(text, size),
                Ascent = face.Ascent(size) + rise,
                Descent = Math.Max(0, face.Descent(size) - rise),
                Url = url,
                Target = target,
            };
        }

        private string ListMarker(DocxList list, int level)
        {
            if (!_counters.TryGetValue(list, out var counters))
            {
                counters = new int[9];
                _counters[list] = counters;
            }

            counters[level]++;
            for (var deeper = level + 1; deeper < 9; deeper++)
            {
                counters[deeper] = 0;
            }

            if (list.Kind == DocxListKind.Bullet)
            {
                return level % 3 == 0 ? "•" : level % 3 == 1 ? "–" : "»";
            }

            var n = counters[level];
            return level % 3 == 0 ? Interlocked(n) + "." : level % 3 == 1 ? Letters(n) + ")" : Roman(n).ToLowerInvariant() + ".";
        }

        private static string Letters(int n)
        {
            var text = string.Empty;
            for (; n > 0; n = (n - 1) / 26)
            {
                text = (char)('a' + ((n - 1) % 26)) + text;
            }

            return text;
        }

        private static string Roman(int n)
        {
            var values = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            var symbols = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            var text = new StringBuilder();
            for (var i = 0; i < values.Length && n > 0; i++)
            {
                while (n >= values[i])
                {
                    text.Append(symbols[i]);
                    n -= values[i];
                }
            }

            return text.ToString();
        }

        private string FieldText(DocxField field, Fields? fields)
        {
            switch (field)
            {
                case DocxField.Page: return fields?.Page ?? "1";
                case DocxField.NumPages: return fields?.Total ?? "1";
                case DocxField.SectionPages: return fields?.SectionPages ?? "1";
                default: return _doc.FieldResult(field);
            }
        }

        private string ReferenceText(DocxRun run, DocxParagraph target)
        {
            switch (run.ReferenceKind)
            {
                case DocxReferenceKind.Label: return target.CaptionLabel ?? string.Empty;
                case DocxReferenceKind.Number: return _numbers.TryGetValue(target, out var number) ? number : string.Empty;
                case DocxReferenceKind.Page: return _placed.TryGetValue(target, out var place) ? PageLabel(place.Page) : "1";
                default: return target.Text;
            }
        }

        /// <summary>
        /// A paragraph broken into lines: one band per line, the first with the space before, the last with the space
        /// after; the lines and the next paragraph kept together as the style asks.
        /// </summary>
        private List<Band> Paragraph(DocxParagraph p, double width, DocxAlignment? fallback, Cell? cell, Fields? fields, List<Band> notes)
        {
            var s = Resolve(p, fallback, cell);
            // Contextual spacing, as list items have: no space between paragraphs of one style.
            if (s.Contextual && _previous != null && _previous.Style == p.Style)
            {
                s.Before = 0;
            }

            if (s.Contextual && _next == p.Style)
            {
                s.After = 0;
            }

            var own = new List<Band>();
            var atoms = Atoms(p, s, fields, own, width);
            var lines = Break(atoms, s, width, p.TabStops);
            var bands = new List<Band>();
            var shaded = p.Shading;
            var border = p.Border;
            for (var i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var last = i == lines.Count - 1;
                var band = new Band
                {
                    Height = line.Height,
                    Before = i == 0 ? s.Before + (border != null ? 4 : 0) : 0,
                    After = last ? s.After + (border != null ? 4 : 0) : 0,
                    KeepNext = (last && s.KeepNext) || (!last && (s.KeepLines || lines.Count <= 3)),
                    PageBreakBefore = (i == 0 && p.PageBreakBefore) || line.PageBreakBefore,
                    Starts = i == 0 ? p : null,
                };
                var first = i == 0;
                band.Draw = (ctx, x, y) =>
                {
                    if (shaded != null || border != null)
                    {
                        double top = y - (first ? (border != null ? 4 : 0) : 0), bottom = y + line.Height + (last ? (border != null ? 4 : 0) : 0);
                        double left = x + s.Left - (border != null ? 4 : 0), right = x + width - s.Right + (border != null ? 4 : 0);
                        if (shaded != null) { ctx.Page.FillRectangle(left, top, right - left, bottom - top, shaded); }
                        if (border is DocxBorder b)
                        {
                            ctx.Page.DrawLine(left, top, left, bottom, b.Color, b.Width).DrawLine(right, top, right, bottom, b.Color, b.Width);
                            if (first) { ctx.Page.DrawLine(left, top, right, top, b.Color, b.Width); }
                            if (last) { ctx.Page.DrawLine(left, bottom, right, bottom, b.Color, b.Width); }
                        }
                    }

                    DrawLine(ctx, line, x, y);
                };
                bands.Add(band);
            }

            if (own.Count > 0 && bands.Count > 0)
            {
                bands[bands.Count - 1].Notes.AddRange(own);
            }

            if (cell == null)
            {
                _previous = p;
            }

            return bands;
        }

        /// <summary>
        /// Breaks atoms into lines that fit the width, then places them: tab stops, alignment, justification.
        /// </summary>
        private List<Line> Break(List<Atom> atoms, Style s, double width, IReadOnlyList<DocxTabStop> stops)
        {
            var lines = new List<Line>();
            var current = new Line();
            double x = s.Left + Math.Max(0, s.First);
            var right = width - s.Right;
            var pageBreak = false;

            void Finish(bool forced)
            {
                while (current.Atoms.Count > 0 && current.Atoms[current.Atoms.Count - 1].Kind == AtomKind.Space)
                {
                    current.Atoms.RemoveAt(current.Atoms.Count - 1);
                }

                current.Last = forced;
                current.PageBreakBefore = pageBreak;
                pageBreak = false;
                lines.Add(current);
                current = new Line();
                x = s.Left;
            }

            for (var i = 0; i < atoms.Count; i++)
            {
                var atom = atoms[i];
                switch (atom.Kind)
                {
                    case AtomKind.Break:
                        Finish(true);
                        continue;
                    case AtomKind.PageBreak:
                        if (current.Atoms.Count > 0) { Finish(true); }
                        pageBreak = true;
                        continue;
                    case AtomKind.Tab:
                        current.Atoms.Add(atom);
                        var stop = NextStop(x, s, width, stops);
                        if (stop.Alignment == DocxTabAlignment.Left)
                        {
                            x = stop.Position;
                        }
                        else
                        {
                            // The text after a right or centred stop ends at (or around) the stop, not after it.
                            var segment = 0.0;
                            for (var j = i + 1; j < atoms.Count && atoms[j].Kind != AtomKind.Tab && atoms[j].Kind != AtomKind.Break; j++)
                            {
                                segment += atoms[j].Width;
                            }

                            x = Math.Max(x, stop.Position - (stop.Alignment == DocxTabAlignment.Right ? segment : segment / 2));
                        }

                        continue;
                    case AtomKind.Space:
                        current.Atoms.Add(atom);
                        x += atom.Width;
                        continue;
                }

                if (x + atom.Width > right + 0.01 && current.Atoms.Exists(a => a.Kind != AtomKind.Space && a.Kind != AtomKind.Tab))
                {
                    Finish(false);
                }

                current.Atoms.Add(atom);
                x += atom.Width;
            }

            Finish(true);
            foreach (var line in lines)
            {
                Place(line, s, width, stops, lines[0] == line);
            }

            return lines;
        }

        private static (double Position, DocxTabAlignment Alignment, DocxTabLeader Leader) NextStop(double x, Style s, double width, IReadOnlyList<DocxTabStop> stops)
        {
            // A hanging indent is a stop of its own, as in Word: the text after a list marker or number starts there.
            if (s.First < 0 && x < s.Left - 0.01)
            {
                return (s.Left, DocxTabAlignment.Left, DocxTabLeader.None);
            }

            (double, DocxTabAlignment, DocxTabLeader)? best = null;
            foreach (var stop in stops)
            {
                var position = stop.Position * Cm;
                if (position > x + 0.01 && (best == null || position < best.Value.Item1))
                {
                    best = (position, stop.Alignment, stop.Leader);
                }
            }

            if (best != null)
            {
                return best.Value;
            }

            const double Default = 35.4;
            return (Math.Floor((x + 0.01) / Default) * Default + Default, DocxTabAlignment.Left, DocxTabLeader.None);
        }

        /// <summary>
        /// The x of every atom of a line, its height and baseline.
        /// </summary>
        private void Place(Line line, Style s, double width, IReadOnlyList<DocxTabStop> stops, bool first)
        {
            double ascent = 0, descent = 0, height = 0;
            foreach (var atom in line.Atoms)
            {
                if (atom.Kind == AtomKind.Word || atom.Kind == AtomKind.Space)
                {
                    ascent = Math.Max(ascent, atom.Ascent);
                    descent = Math.Max(descent, atom.Descent);
                    height = Math.Max(height, atom.Font!.LineHeight(atom.Size + (atom.Rise != 0 ? atom.Size * 0.5 : 0)) * s.Line);
                }
                else if (atom.Kind == AtomKind.Image || atom.Kind == AtomKind.Box)
                {
                    ascent = Math.Max(ascent, atom.Ascent);
                }
            }

            if (height == 0 && ascent == 0)
            {
                var font = Font(s.Family, s.Bold, s.Italic);
                ascent = font.Ascent(s.Size);
                descent = font.Descent(s.Size);
                height = font.LineHeight(s.Size) * s.Line;
            }

            line.Baseline = ascent + Math.Max(0, (height - ascent - descent) * 0.25);
            line.Height = Math.Max(height, line.Baseline + descent);

            var start = s.Left + (first ? s.First : 0);
            var hasTabs = line.Atoms.Exists(a => a.Kind == AtomKind.Tab);
            var natural = 0.0;
            var spaces = 0;
            foreach (var atom in line.Atoms)
            {
                natural += atom.Kind == AtomKind.Tab ? 0 : atom.Width;
                if (atom.Kind == AtomKind.Space) { spaces++; }
            }

            var room = width - s.Right - start;
            var offset = 0.0;
            var stretch = 0.0;
            if (!hasTabs)
            {
                if (s.Align == DocxAlignment.Center) { offset = (room - natural) / 2; }
                else if (s.Align == DocxAlignment.Right) { offset = room - natural; }
                else if (s.Align == DocxAlignment.Justify && !line.Last && spaces > 0) { stretch = (room - natural) / spaces; }
            }

            var x = start + Math.Max(0, offset);
            for (var i = 0; i < line.Atoms.Count; i++)
            {
                var atom = line.Atoms[i];
                if (atom.Kind == AtomKind.Tab)
                {
                    var (position, alignment, leader) = NextStop(x, s, width, stops);
                    var segment = 0.0;
                    for (var j = i + 1; j < line.Atoms.Count && line.Atoms[j].Kind != AtomKind.Tab; j++)
                    {
                        segment += line.Atoms[j].Width;
                    }

                    var target = alignment == DocxTabAlignment.Right ? position - segment : alignment == DocxTabAlignment.Center ? position - (segment / 2) : position;
                    target = Math.Max(target, x);
                    atom.X = x;
                    atom.Width = target - x;
                    atom.Leader = leader;
                    x = target;
                    continue;
                }

                atom.X = x;
                x += atom.Width + (atom.Kind == AtomKind.Space ? stretch : 0);
            }
        }

        /// <summary>
        /// Draws a line's atoms, joining neighbours of one font into one text so the file stays small.
        /// </summary>
        private void DrawLine(Context ctx, Line line, double x, double y)
        {
            var page = ctx.Page;
            var baseline = y + line.Baseline;
            var text = new StringBuilder();
            Atom? run = null;
            var runX = 0.0;

            void Flush()
            {
                if (run != null && text.Length > 0)
                {
                    page.DrawText(text.ToString(), x + runX, baseline - run.Rise, run.Font!, run.Size, run.Color ?? "#000000");
                }

                text.Clear();
                run = null;
            }

            foreach (var atom in line.Atoms)
            {
                if (atom.Kind == AtomKind.Word || atom.Kind == AtomKind.Space)
                {
                    if (atom.Back != null)
                    {
                        page.FillRectangle(x + atom.X, baseline - atom.Font!.Ascent(atom.Size), atom.Width, atom.Font.Ascent(atom.Size) + atom.Font.Descent(atom.Size), atom.Back);
                    }

                    var joins = run != null && ReferenceEquals(run.Font, atom.Font) && run.Size == atom.Size && run.Color == atom.Color && run.Rise == atom.Rise
                        && Math.Abs(runX + run.Font!.Measure(text.ToString(), run.Size) - atom.X) < 0.01;
                    if (!joins)
                    {
                        Flush();
                        run = atom;
                        runX = atom.X;
                    }

                    text.Append(atom.Text);
                    if (atom.Underline || atom.Strike)
                    {
                        var thickness = Math.Max(0.4, atom.Size * 0.06);
                        var at = atom.Underline ? baseline - atom.Rise + (atom.Size * 0.12) : baseline - atom.Rise - (atom.Size * 0.28);
                        page.DrawLine(x + atom.X, at, x + atom.X + atom.Width, at, atom.Color ?? "#000000", thickness);
                    }

                    if (atom.Url != null || atom.Target != null)
                    {
                        ctx.Links.Add((page, x + atom.X, baseline - atom.Ascent, atom.Width, atom.Ascent + atom.Descent, atom.Url, atom.Target));
                    }

                    continue;
                }

                Flush();
                if (atom.Kind == AtomKind.Tab && atom.Leader != DocxTabLeader.None && atom.Width > 6)
                {
                    // Leader dots or a line up to the text after the tab.
                    var dash = atom.Leader == DocxTabLeader.Dot ? PdfDash.Dotted : atom.Leader == DocxTabLeader.Hyphen ? PdfDash.Dashed : PdfDash.Solid;
                    page.DrawLine(x + atom.X + 2, baseline, x + atom.X + atom.Width - 2, baseline, "#7F7F7F", 0.8, dash);
                }
                else if (atom.Kind == AtomKind.Image)
                {
                    page.DrawImage(atom.Image!, x + atom.X, baseline - atom.Ascent, atom.Width, atom.Ascent);
                }
                else if (atom.Kind == AtomKind.Box)
                {
                    var side = atom.Size * 0.72;
                    double left = x + atom.X + (atom.Size * 0.04), top = baseline - side;
                    page.StrokeRectangle(left, top, side, side, atom.Color ?? "#000000", Math.Max(0.5, atom.Size * 0.06));
                    if (atom.Checked)
                    {
                        page.DrawLine(left + (side * 0.2), top + (side * 0.2), left + (side * 0.8), top + (side * 0.8), atom.Color ?? "#000000", Math.Max(0.6, atom.Size * 0.07));
                        page.DrawLine(left + (side * 0.8), top + (side * 0.2), left + (side * 0.2), top + (side * 0.8), atom.Color ?? "#000000", Math.Max(0.6, atom.Size * 0.07));
                    }
                }
            }

            Flush();
        }

        // ---------------------------------------------------------------------------------------------- blocks and tables

        /// <summary>
        /// The bands of a container's blocks: paragraphs, tables, contents, bibliography.
        /// </summary>
        private List<Band> Blocks(DocxContainer container, double width, DocxAlignment? fallback, Cell? cell, Fields? fields, List<Band> notes)
        {
            var bands = new List<Band>();
            var blocks = container.Blocks;
            for (var b = 0; b < blocks.Count; b++)
            {
                var block = blocks[b];
                switch (block)
                {
                    case DocxParagraph p:
                        _next = b + 1 < blocks.Count && blocks[b + 1] is DocxParagraph following ? following.Style : null;
                        bands.AddRange(Paragraph(p, width, fallback, cell, fields, notes));
                        _next = null;
                        break;
                    case DocxTable table:
                        bands.AddRange(Table(table, width, fields, notes));
                        _previous = null;
                        break;
                    case DocxTableOfContents toc:
                        foreach (var entry in Contents(toc, width))
                        {
                            bands.AddRange(Paragraph(entry, width, null, null, fields, notes));
                        }

                        break;
                    case DocxBibliography _:
                        foreach (var source in _doc._sources)
                        {
                            var entry = new DocxParagraph(BibliographyStyleId).Add(_doc.EntryText(source));
                            bands.AddRange(Paragraph(entry, width, null, null, fields, notes));
                        }

                        break;
                }
            }

            return bands;
        }

        /// <summary>
        /// The entries of a table of contents or of a list of figures: text, dot leaders, the page, a link.
        /// </summary>
        private IEnumerable<DocxParagraph> Contents(DocxTableOfContents toc, double width)
        {
            var stop = (width / Cm) - 0.01;
            if (toc.Captions is DocxCaptionKind kind)
            {
                foreach (var section in _doc._sections)
                {
                    foreach (var block in section.Blocks)
                    {
                        if (block is DocxParagraph p && p.Caption == kind)
                        {
                            yield return Entry(FiguresStyle, p.CaptionText ?? p.Text, p, stop);
                        }
                    }
                }

                yield break;
            }

            foreach (var (paragraph, level, number) in _headings)
            {
                if (level <= toc.Levels)
                {
                    yield return Entry("TOC" + Interlocked(level), (number.Length > 0 ? number + "  " : string.Empty) + paragraph.Text, paragraph, stop);
                }
            }
        }

        private DocxParagraph Entry(string style, string text, DocxParagraph target, double stop)
        {
            var entry = new DocxParagraph(style);
            entry.AddTabStop(stop, DocxTabAlignment.Right, DocxTabLeader.Dot);
            entry.Add(text).AddTab().Add(_placed.TryGetValue(target, out var place) ? PageLabel(place.Page) : "0");
            entry.ContentsTarget = target;
            return entry;
        }

        /// <summary>
        /// A table: one band per row, each drawing its cells with their fills, lines and content; header rows marked
        /// to repeat, the rows of a row span kept together.
        /// </summary>
        private List<Band> Table(DocxTable table, double width, Fields? fields, List<Band> notes)
        {
            var columns = table.Columns;
            var widths = new double[columns];
            for (var c = 0; c < columns; c++)
            {
                widths[c] = table.ColumnWidths != null ? table.ColumnWidths[c] * Cm : width / columns;
            }

            var total = 0.0;
            foreach (var w in widths) { total += w; }
            var left = table.Centered ? Math.Max(0, (width - total) / 2) : 0;
            var lines = table.Look != DocxTableLook.Plain;
            var banded = table.Look == DocxTableLook.Banded;
            const double PadX = 5.4, PadY = 2.85;

            var rows = new List<(DocxTableRow Row, List<(int Column, int Span, int RowSpan, DocxTableCell Cell, List<Band> Content, double Height)> Cells, double Height)>();
            var covered = new int[columns];
            for (var r = 0; r < table.Rows.Count; r++)
            {
                var row = table.Rows[r];
                var header = r == 0 && lines;
                var cells = new List<(int, int, int, DocxTableCell, List<Band>, double)>();
                var height = row.Height is double minimum ? minimum * Cm : 0;
                var c = 0;
                while (c < columns)
                {
                    if (covered[c] > 0)
                    {
                        covered[c]--;
                        c++;
                        continue;
                    }

                    var cell = row.Cells[c];
                    var span = Math.Max(1, Math.Min(cell.ColumnSpan, columns - c));
                    var inner = -2 * PadX;
                    for (var k = c; k < c + span; k++) { inner += widths[k]; }
                    var look = new Cell { Header = header, HeaderColor = banded ? Look.TableHeaderColor : Look.BodyColor };
                    var saved = _previous;
                    var content = Blocks(cell, Math.Max(1, inner), table.AlignmentOf(c), look, fields, notes);
                    _previous = saved;
                    var contentHeight = (2 * PadY) + Stack(content);
                    if (cell.RowSpan > 1)
                    {
                        for (var k = c; k < c + span; k++) { covered[k] = cell.RowSpan - 1; }
                    }
                    else
                    {
                        height = Math.Max(height, contentHeight);
                    }

                    cells.Add((c, span, Math.Max(1, cell.RowSpan), cell, content, contentHeight));
                    for (var k = c + 1; k < c + span; k++) { covered[k] = Math.Max(covered[k], 0); }
                    c += span;
                }

                rows.Add((row, cells, Math.Max(height, 12)));
            }

            // A row span taller than its rows lengthens the last of them.
            var heights = new double[rows.Count];
            for (var r = 0; r < rows.Count; r++) { heights[r] = rows[r].Height; }
            for (var r = 0; r < rows.Count; r++)
            {
                foreach (var cell in rows[r].Cells)
                {
                    if (cell.RowSpan > 1)
                    {
                        var last = Math.Min(rows.Count - 1, r + cell.RowSpan - 1);
                        var sum = 0.0;
                        for (var k = r; k <= last; k++) { sum += heights[k]; }
                        if (cell.Height > sum) { heights[last] += cell.Height - sum; }
                    }
                }
            }

            var bands = new List<Band>();
            var headers = new List<Band>();
            var spanUntil = -1;
            for (var r = 0; r < rows.Count; r++)
            {
                var (row, cells, _) = rows[r];
                var index = r;
                var header = r == 0 && lines;
                var bandFill = banded && !header && (r % 2 == 1);
                foreach (var cell in cells)
                {
                    spanUntil = Math.Max(spanUntil, r + cell.RowSpan - 1);
                }

                var band = new Band { Height = heights[r], KeepNext = r < spanUntil || row.IsHeader, IsHeader = row.IsHeader, Headers = headers };
                band.Draw = (ctx, x, y) =>
                {
                    foreach (var (column, span, rowSpan, cell, content, _) in cells)
                    {
                        var cx = x + left;
                        for (var k = 0; k < column; k++) { cx += widths[k]; }
                        var cw = 0.0;
                        for (var k = column; k < column + span; k++) { cw += widths[k]; }
                        var ch = 0.0;
                        for (var k = index; k < Math.Min(heights.Length, index + rowSpan); k++) { ch += heights[k]; }
                        var fill = cell.Shading ?? (header ? (banded ? Look.TableHeaderFill : Look.TableBandFill) : bandFill ? Look.TableBandFill : null);
                        if (fill != null) { ctx.Page.FillRectangle(cx, y, cw, ch, fill); }
                        if (lines) { ctx.Page.StrokeRectangle(cx, y, cw, ch, Look.TableBorderColor, 0.5); }
                        var used = Stack(content);
                        var top = y + PadY;
                        if (cell.VerticalAlignment == DocxCellAlignment.Center) { top += Math.Max(0, (ch - (2 * PadY) - used) / 2); }
                        else if (cell.VerticalAlignment == DocxCellAlignment.Bottom) { top += Math.Max(0, ch - (2 * PadY) - used); }
                        Paint(ctx, content, cx + PadX, top);
                    }
                };
                if (row.IsHeader && r == headers.Count)
                {
                    headers.Add(band);
                }

                bands.Add(band);
            }

            if (bands.Count > 0)
            {
                bands[0].Before = 0;
                bands[bands.Count - 1].After = 6;
            }

            return bands;
        }

        /// <summary>
        /// The height of bands stacked without page breaks, their spacing included.
        /// </summary>
        private static double Stack(List<Band> bands)
        {
            var height = 0.0;
            for (var i = 0; i < bands.Count; i++)
            {
                height += (i == 0 ? 0 : bands[i].Before) + bands[i].Height + (i == bands.Count - 1 ? 0 : bands[i].After);
            }

            return height;
        }

        private static void Paint(Context ctx, List<Band> bands, double x, double y)
        {
            for (var i = 0; i < bands.Count; i++)
            {
                y += i == 0 ? 0 : bands[i].Before;
                bands[i].Draw!(ctx, x, y);
                y += bands[i].Height + bands[i].After;
            }
        }

        // ---------------------------------------------------------------------------------------------- pages

        /// <summary>
        /// Fills pages with the bands of every section: a new page per section, page breaks where asked, keeping
        /// bands with the next as asked, header rows of a table again at the top of a page it continues on, and the
        /// footnotes of the lines on a page at its foot.
        /// </summary>
        private List<PageLayout> Flow()
        {
            _counters = new Dictionary<DocxList, int[]>();
            _footnotes = 0;
            var pages = new List<PageLayout>();
            var number = 0;
            for (var si = 0; si < _doc._sections.Count; si++)
            {
                var section = _doc._sections[si];
                var (paperWidth, paperHeight) = section.PaperTwips;
                var margins = section.Margins;
                var width = section.TextWidthTwips / 20.0;
                double top = margins.Top * Cm, bottom = (paperHeight / 20.0) - (margins.Bottom * Cm);
                number = section.PageNumberStart is int start ? start - 1 : number;
                var first = pages.Count;
                _previous = null;
                var notes = new List<Band>();
                var bands = section.Source != null ? new List<Band>() : Blocks(section, width, null, null, null, notes);

                PageLayout NewPage()
                {
                    var page = new PageLayout { Section = section, SectionIndex = si, Number = ++number, Width = paperWidth / 20.0, Height = paperHeight / 20.0, Top = top, Bottom = bottom, Left = margins.Left * Cm, TextWidth = width, Cursor = top };
                    pages.Add(page);
                    return page;
                }

                var current = NewPage();
                for (var i = 0; i < bands.Count; i++)
                {
                    var band = bands[i];
                    if (band.PageBreakBefore && current.Bands.Count > 0)
                    {
                        current = NewPage();
                        current.Natural = false;
                    }

                    var before = current.Bands.Count == 0 && current.Natural ? 0 : band.Before;
                    var noteHeight = band.Notes.Count > 0 ? Stack(band.Notes) + 8 : 0;
                    var chain = band.Height;
                    for (var k = i; k < bands.Count - 1 && bands[k].KeepNext && chain < bottom - top; k++)
                    {
                        chain += bands[k].After + bands[k + 1].Before + bands[k + 1].Height;
                    }

                    var limit = current.Bottom - current.NoteHeight - (current.NoteHeight > 0 || noteHeight > 0 ? noteHeight : 0);
                    if (current.Bands.Count > 0 && (current.Cursor + before + chain > limit))
                    {
                        current = NewPage();
                        before = 0;
                        if (band.Headers != null && !band.IsHeader && band.Headers.Count > 0)
                        {
                            foreach (var repeat in band.Headers)
                            {
                                current.Bands.Add((repeat, current.Cursor));
                                current.Cursor += repeat.Height;
                            }
                        }
                    }

                    current.Cursor += before;
                    current.Bands.Add((band, current.Cursor));
                    if (band.Starts != null)
                    {
                        current.Starts.Add((band.Starts, current.Cursor));
                    }

                    current.Cursor += band.Height + band.After;
                    if (band.Notes.Count > 0)
                    {
                        current.Notes.AddRange(band.Notes);
                        current.NoteHeight = Stack(current.Notes) + 8;
                    }
                }

                for (var p = first; p < pages.Count; p++)
                {
                    pages[p].SectionPages = pages.Count - first;
                }
            }

            return pages;
        }

        private Dictionary<DocxParagraph, (int Page, double Y)> Placements(List<PageLayout> pages)
        {
            var placed = new Dictionary<DocxParagraph, (int, double)>();
            for (var i = 0; i < pages.Count; i++)
            {
                foreach (var (paragraph, y) in pages[i].Starts)
                {
                    if (!placed.ContainsKey(paragraph)) { placed[paragraph] = (i, y); }
                }
            }

            _numbersOfPages = new string[pages.Count];
            for (var i = 0; i < pages.Count; i++) { _numbersOfPages[i] = Format(pages[i]); }
            return placed;
        }

        private string[] _numbersOfPages = Array.Empty<string>();

        private string PageLabel(int page) => page < _numbersOfPages.Length ? _numbersOfPages[page] : Interlocked(page + 1);

        private static string Format(PageLayout page) =>
            page.Section.PageNumberFormat == DocxNumberFormat.LowerRoman ? Roman(page.Number).ToLowerInvariant()
            : page.Section.PageNumberFormat == DocxNumberFormat.UpperRoman ? Roman(page.Number) : Interlocked(page.Number);

        /// <summary>
        /// Draws the pages: watermark, header, footer, the bands (centred between the margins where the section asks),
        /// footnotes; then the links and bookmarks, once every page exists.
        /// </summary>
        private void Emit(List<PageLayout> pages)
        {
            _numbersOfPages = new string[pages.Count];
            for (var i = 0; i < pages.Count; i++) { _numbersOfPages[i] = Format(pages[i]); }
            var even = false;
            foreach (var section in _doc._sections)
            {
                even |= section.EvenHeaderPart != null || section.EvenFooterPart != null;
            }

            var links = new List<(PdfPage Page, double X, double Y, double Width, double Height, string? Url, DocxParagraph? Target)>();
            var pdfPages = new List<PdfPage>();
            for (var i = 0; i < pages.Count; i++)
            {
                var layout = pages[i];
                var page = _pdf.AddPage(layout.Width, layout.Height);
                pdfPages.Add(page);
                var ctx = new Context(page, links);
                var fields = new Fields(_numbersOfPages[i], Interlocked(pages.Count), Interlocked(layout.SectionPages));
                var firstOfSection = i == 0 || pages[i - 1].SectionIndex != layout.SectionIndex;
                var section = layout.Section;

                if (section.Watermark != null) { Watermark(page, layout, section); }

                var header = Part(layout.SectionIndex, firstOfSection, even && layout.Number % 2 == 0, false);
                if (header != null)
                {
                    var bands = Blocks(header, layout.TextWidth, null, null, fields, new List<Band>());
                    Paint(ctx, bands, layout.Left, section.Margins.Header * Cm);
                }

                var footer = Part(layout.SectionIndex, firstOfSection, even && layout.Number % 2 == 0, true);
                if (footer != null)
                {
                    var bands = Blocks(footer, layout.TextWidth, null, null, fields, new List<Band>());
                    Paint(ctx, bands, layout.Left, layout.Height - (section.Margins.Footer * Cm) - Stack(bands));
                }

                var shift = 0.0;
                if (section.VerticalAlignment != DocxVerticalAlignment.Top && layout.Bands.Count > 0)
                {
                    var used = layout.Cursor - layout.Top;
                    var free = Math.Max(0, layout.Bottom - layout.Top - used);
                    shift = section.VerticalAlignment == DocxVerticalAlignment.Center ? free / 2 : free;
                }

                foreach (var (band, y) in layout.Bands)
                {
                    band.Draw!(ctx, layout.Left, y + shift);
                }

                if (layout.Notes.Count > 0)
                {
                    var top = layout.Bottom - Stack(layout.Notes);
                    page.DrawLine(layout.Left, top - 4, layout.Left + 120, top - 4, "#7F7F7F", 0.5);
                    Paint(ctx, layout.Notes, layout.Left, top);
                }

                foreach (var (paragraph, y) in layout.Starts)
                {
                    _placed[paragraph] = (i, y + shift);
                }
            }

            foreach (var (page, x, y, w, h, url, target) in links)
            {
                if (url != null)
                {
                    page.AddLink(x, y, w, h, url);
                }
                else if (target != null && _placed.TryGetValue(target, out var place))
                {
                    page.AddLink(x, y, w, h, pdfPages[place.Page], Math.Max(0, place.Y - 6));
                }
            }

            foreach (var (paragraph, level, number) in _headings)
            {
                if (_placed.TryGetValue(paragraph, out var place))
                {
                    _pdf.AddBookmark((number.Length > 0 ? number + " " : string.Empty) + paragraph.Text, pdfPages[place.Page], Math.Max(0, place.Y - 6), level - 1);
                }
            }
        }

        /// <summary>
        /// The header or footer of a page: the first-page one on a section's first page when it has one, the even one
        /// on even pages, else the default; a section without its own takes the previous section's.
        /// </summary>
        private DocxHeaderFooter? Part(int sectionIndex, bool firstPage, bool evenPage, bool footer)
        {
            var section = _doc._sections[sectionIndex];
            if (firstPage && section.DifferentFirstPage)
            {
                return footer ? section.FirstFooterPart : section.FirstHeaderPart;
            }

            for (var i = sectionIndex; i >= 0; i--)
            {
                var s = _doc._sections[i];
                var part = evenPage ? (footer ? s.EvenFooterPart : s.EvenHeaderPart) : (footer ? s.DefaultFooterPart : s.DefaultHeaderPart);
                if (part != null)
                {
                    return part;
                }
            }

            return null;
        }

        /// <summary>
        /// The watermark: the text slanted up from bottom left to top right through the middle of the page, as
        /// large as the text width allows, in the colour half towards white.
        /// </summary>
        private void Watermark(PdfPage page, PageLayout layout, DocxSection section)
        {
            var font = Font(Look.BodyFont, true, false);
            var text = section.Watermark!;
            var unit = font.Measure(text, 1);
            if (unit <= 0)
            {
                return;
            }

            var size = layout.TextWidth * 0.9 / unit;
            var rgb = Divit.Packaging.OfficeText.Rgb(section.WatermarkColor);
            var color = new StringBuilder("#");
            for (var i = 0; i < 6; i += 2)
            {
                var value = Convert.ToInt32(rgb.Substring(i, 2), 16);
                color.Append(((value + 255) / 2).ToString("X2", CultureInfo.InvariantCulture));
            }

            double cx = layout.Width / 2, cy = layout.Height / 2, angle = 45 * Math.PI / 180;
            var half = unit * size / 2;
            var lift = font.CapHeight(size) / 2;
            var x = cx - (half * Math.Cos(angle)) + (lift * Math.Sin(angle));
            var y = cy + (half * Math.Sin(angle)) + (lift * Math.Cos(angle));
            page.DrawTurnedText(text, x, y, font, size, color.ToString(), 45);
        }
    }

    private sealed class Style
    {
        public string Family = string.Empty;
        public double Size;
        public bool Bold;
        public bool Italic;
        public string Color = "#000000";
        public double Before;
        public double After;
        public double Line = 1;
        public double Left;
        public double Right;
        public double First;
        public DocxAlignment Align;
        public bool KeepNext;
        public bool KeepLines;
        public bool Contextual;
    }

    /// <summary>
    /// What a table cell gives its paragraphs: whether they are in the header row, and its text colour.
    /// </summary>
    private sealed class Cell
    {
        public bool Header;
        public string HeaderColor = "#000000";
    }

    /// <summary>
    /// The values of page fields on the page being drawn.
    /// </summary>
    private sealed class Fields
    {
        public Fields(string page, string total, string sectionPages)
        {
            Page = page;
            Total = total;
            SectionPages = sectionPages;
        }

        public string Page { get; }

        public string Total { get; }

        public string SectionPages { get; }
    }

    private enum AtomKind
    {
        Word,
        Space,
        Tab,
        Break,
        PageBreak,
        Image,
        Box,
    }

    private sealed class Atom
    {
        public AtomKind Kind;
        public string Text = string.Empty;
        public PdfFont? Font;
        public double Size;
        public string? Color;
        public bool Underline;
        public bool Strike;
        public string? Back;
        public double Rise;
        public double Width;
        public double Ascent;
        public double Descent;
        public double X;
        public byte[]? Image;
        public bool Checked;
        public string? Url;
        public DocxParagraph? Target;
        public DocxTabLeader Leader;
    }

    private sealed class Line
    {
        public List<Atom> Atoms { get; } = new List<Atom>();

        public double Height;
        public double Baseline;
        public bool Last;
        public bool PageBreakBefore;
    }

    /// <summary>
    /// A strip of a page that is never split: a line of a paragraph, a table row.
    /// </summary>
    private sealed class Band
    {
        public double Height;
        public double Before;
        public double After;
        public bool KeepNext;
        public bool PageBreakBefore;
        public bool IsHeader;

        /// <summary>
        /// The header rows of the table the band is a row of, repeated when the table goes on to a new page.
        /// </summary>
        public List<Band>? Headers;

        /// <summary>
        /// The paragraph whose first line this is, for the page numbers of contents and references.
        /// </summary>
        public DocxParagraph? Starts;

        public List<Band> Notes { get; } = new List<Band>();

        public Action<Context, double, double>? Draw;
    }

    private sealed class Context
    {
        public Context(PdfPage page, List<(PdfPage, double, double, double, double, string?, DocxParagraph?)> links)
        {
            Page = page;
            Links = links;
        }

        public PdfPage Page { get; }

        public List<(PdfPage Page, double X, double Y, double Width, double Height, string? Url, DocxParagraph? Target)> Links { get; }
    }

    private sealed class PageLayout
    {
        public DocxSection Section = null!;
        public int SectionIndex;
        public int Number;
        public int SectionPages;
        public double Width;
        public double Height;
        public double Top;
        public double Bottom;
        public double Left;
        public double TextWidth;
        public double Cursor;
        public double NoteHeight;
        public bool Natural = true;

        public List<(Band Band, double Y)> Bands { get; } = new List<(Band, double)>();

        public List<(DocxParagraph Paragraph, double Y)> Starts { get; } = new List<(DocxParagraph, double)>();

        public List<Band> Notes { get; } = new List<Band>();
    }
}
