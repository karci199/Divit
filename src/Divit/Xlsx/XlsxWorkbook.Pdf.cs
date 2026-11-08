#nullable enable
namespace Divit.Xlsx;

using Divit.Pdf;
using System;
using System.Collections.Generic;
using System.Globalization;

/// <content>
/// The workbook printed as PDF: each visible sheet over pages as Excel prints it.
/// </content>
public sealed partial class XlsxWorkbook
{
    /// <summary>
    /// Prints the visible sheets as PDF pages, as Excel would: paper, orientation and margins of each sheet, fitted to
    /// its width (or height) when asked, pages down then across, title rows repeated, manual page breaks, header and
    /// footer. Cells keep their values in their number formats, fonts, fills, borders, alignment, wrapping and
    /// merges; charts and images stand where they are anchored. More can be drawn on the pages before saving.
    /// </summary>
    /// <remarks>
    /// Streamed sheets are left out (their rows exist only while the workbook is saved), as are conditional formats,
    /// notes and formula results that were not given.
    /// </remarks>
    /// <returns>The PDF document.</returns>
    public PdfDocument ToPdf()
    {
        var pdf = new PdfDocument { Title = Title, Author = Author, Subject = Subject, Application = "Divit", Created = Created };
        foreach (var sheet in _sheets)
        {
            if (sheet.Visibility == XlsxSheetVisibility.Visible && !sheet.IsStreamed)
            {
                new SheetPrinter(this, sheet, pdf).Print();
            }
        }

        if (pdf.Pages.Count == 0)
        {
            pdf.AddPage();
        }

        return pdf;
    }

    /// <summary>
    /// Saves the workbook printed as PDF (<see cref="ToPdf"/>).
    /// </summary>
    /// <param name="path">Target file.</param>
    public void SavePdf(string path) => ToPdf().Save(path);

    /// <summary>
    /// Prints one sheet: measures its used range, splits it into pages, and draws each page.
    /// </summary>
    private sealed class SheetPrinter
    {
        private const double PadX = 2.25;

        private readonly XlsxWorkbook _book;
        private readonly XlsxSheet _sheet;
        private readonly PdfDocument _pdf;
        private readonly Dictionary<(string Family, bool Bold, bool Italic), PdfFont> _fonts = new Dictionary<(string, bool, bool), PdfFont>();
        private readonly Dictionary<long, (int Row2, int Col2)> _merged = new Dictionary<long, (int, int)>();
        private readonly HashSet<long> _covered = new HashSet<long>();
        private double _scale = 1;

        public SheetPrinter(XlsxWorkbook book, XlsxSheet sheet, PdfDocument pdf)
        {
            _book = book;
            _sheet = sheet;
            _pdf = pdf;
            foreach (var (row1, col1, row2, col2) in sheet.Merges)
            {
                _merged[XlsxSheet.Key(row1, col1)] = (row2, col2);
                for (var r = row1; r <= row2; r++)
                {
                    for (var c = col1; c <= col2; c++)
                    {
                        if (r != row1 || c != col1) { _covered.Add(XlsxSheet.Key(r, c)); }
                    }
                }
            }
        }

        public void Print()
        {
            var sheet = _sheet;
            var (firstRow, firstColumn, lastRow, lastColumn) = Range();
            if (lastRow < firstRow || lastColumn < firstColumn)
            {
                return;
            }

            // The paper, the margins and the printable box in points.
            var (paperWidth, paperHeight) = Paper(sheet.PaperSize);
            if (sheet.Landscape) { (paperWidth, paperHeight) = (paperHeight, paperWidth); }
            var m = sheet.Margins;
            double left = m.Left * 72, top = m.Top * 72, width = paperWidth - ((m.Left + m.Right) * 72), height = paperHeight - ((m.Top + m.Bottom) * 72);

            var titles = sheet.PrintTitleRows;
            var titleHeight = 0.0;
            if (titles is (int t1, int t2))
            {
                for (var r = t1; r <= t2; r++) { titleHeight += RowHeight(r); }
            }

            // Fit to the pages wide (and high) the sheet asks for, never enlarging.
            var totalWidth = 0.0;
            for (var c = firstColumn; c <= lastColumn; c++) { totalWidth += ColumnWidth(c); }
            var totalHeight = 0.0;
            for (var r = firstRow; r <= lastRow; r++) { totalHeight += RowHeight(r); }
            if (sheet.FitsToPage)
            {
                if (sheet.FitWidth > 0) { _scale = Math.Min(_scale, sheet.FitWidth * width / Math.Max(1, totalWidth)); }
                if (sheet.FitHeight > 0) { _scale = Math.Min(_scale, sheet.FitHeight * height / Math.Max(1, totalHeight + titleHeight)); }
            }

            var columnPages = Split(firstColumn, lastColumn, width / _scale, ColumnWidth, sheet.ColumnBreaks);
            var rowPages = Split(firstRow, lastRow, (height / _scale) - titleHeight, RowHeight, sheet.RowBreaks);

            var total = columnPages.Count * rowPages.Count;
            var number = 0;
            foreach (var (c1, c2) in columnPages)
            {
                foreach (var (r1, r2) in rowPages)
                {
                    number++;
                    var page = _pdf.AddPage(paperWidth, paperHeight);
                    var x0 = left;
                    var y0 = top;
                    var pageWidth = 0.0;
                    for (var c = c1; c <= c2; c++) { pageWidth += ColumnWidth(c) * _scale; }
                    var pageHeight = titleHeight * _scale;
                    for (var r = r1; r <= r2; r++) { pageHeight += RowHeight(r) * _scale; }
                    if (sheet.PrintCenteredHorizontally) { x0 += Math.Max(0, (width - pageWidth) / 2); }
                    if (sheet.PrintCenteredVertically) { y0 += Math.Max(0, (height - pageHeight) / 2); }

                    var y = y0;
                    if (titles is (int tf, int tl) && r1 > tl)
                    {
                        // Title rows above every page that does not start with them.
                        DrawRows(page, tf, tl, c1, c2, x0, y);
                        y += titleHeight * _scale;
                    }

                    DrawRows(page, r1, r2, c1, c2, x0, y);
                    DrawObjects(page, r1, r2, c1, c2, x0, y);
                    HeaderFooter(page, number, total, paperWidth, paperHeight);
                }
            }
        }

        // ---------------------------------------------------------------------------------------------- measures

        /// <summary>
        /// The print area, or the used range: cells with a value or style, merges, charts and images.
        /// </summary>
        private (int, int, int, int) Range()
        {
            var sheet = _sheet;
            if (sheet.PrintArea is (int pr1, int pc1, int pr2, int pc2))
            {
                return (pr1, pc1, pr2, pc2);
            }

            int r1 = int.MaxValue, c1 = int.MaxValue, r2 = 0, c2 = 0;
            void Take(int r, int c)
            {
                r1 = Math.Min(r1, r); c1 = Math.Min(c1, c); r2 = Math.Max(r2, r); c2 = Math.Max(c2, c);
            }

            foreach (var key in sheet.Cells.Keys) { Take(XlsxSheet.RowOf(key), XlsxSheet.ColumnOf(key)); }
            foreach (var (a, b, c, d) in sheet.Merges) { Take(a, b); Take(c, d); }
            foreach (var chart in sheet.Charts) { Take(chart.Row, chart.Column); Take(RowAt(chart.Row, chart.Height * 0.75), ColumnAt(chart.Column, chart.Width * 0.75)); }
            foreach (var image in sheet.Images) { Take(image.Row, image.Column); Take(Math.Max(image.LastRow, RowAt(image.Row, image.Height * 0.75)), Math.Max(image.LastColumn, ColumnAt(image.Column, image.Width * 0.75))); }
            // Excel prints from A1, as the sheet looks.
            return r2 == 0 ? (1, 1, 0, 0) : (1, 1, r2, c2);
        }

        /// <summary>
        /// Width of a column in points: Excel's width in characters of the default font (7 pixels each, plus 5), 0 when
        /// hidden.
        /// </summary>
        private double ColumnWidth(int column)
        {
            if (_sheet.Columns.TryGetValue(column, out var line))
            {
                if (line.Hidden) { return 0; }
                if (line.Size is double size) { return Math.Round((size * 7) + 5) * 0.75; }
            }

            return Math.Round((8.43 * 7) + 5) * 0.75;
        }

        private double RowHeight(int row)
        {
            if (_sheet.Rows.TryGetValue(row, out var line))
            {
                if (line.Hidden) { return 0; }
                if (line.Size is double size) { return size; }
            }

            return _book.DefaultRowHeight;
        }

        private int RowAt(int row, double points)
        {
            while (points > 0 && row < XlsxSheet.MaxRow) { points -= RowHeight(row); if (points > 0) { row++; } }
            return row;
        }

        private int ColumnAt(int column, double points)
        {
            while (points > 0 && column < XlsxSheet.MaxColumn) { points -= ColumnWidth(column); if (points > 0) { column++; } }
            return column;
        }

        /// <summary>
        /// Splits rows or columns into pages that fit, at manual breaks too; the title rows are not printed twice.
        /// </summary>
        private static List<(int First, int Last)> Split(int first, int last, double room, Func<int, double> size, IEnumerable<int> breaks)
        {
            var manual = new HashSet<int>(breaks);
            var pages = new List<(int, int)>();
            var start = first;
            var used = 0.0;
            for (var i = first; i <= last; i++)
            {
                var s = size(i);
                if (i > start && used + s > room + 0.01)
                {
                    pages.Add((start, i - 1));
                    start = i;
                    used = 0;
                }

                used += s;
                if (manual.Contains(i) && i < last)
                {
                    pages.Add((start, i));
                    start = i + 1;
                    used = 0;
                }
            }

            if (start <= last) { pages.Add((start, last)); }
            return pages;
        }

        private static (double, double) Paper(XlsxPaperSize size) => size switch
        {
            XlsxPaperSize.Letter => (612, 792),
            XlsxPaperSize.Legal => (612, 1008),
            XlsxPaperSize.A3 => (841.89, 1190.55),
            XlsxPaperSize.A5 => (419.53, 595.28),
            XlsxPaperSize.B5 => (498.9, 708.66),
            _ => (595.28, 841.89),
        };

        private PdfFont Font(string? family, bool bold, bool italic)
        {
            var name = family ?? _book.DefaultFontName;
            if (_fonts.TryGetValue((name, bold, italic), out var font)) { return font; }
            foreach (var candidate in new[] { name, _book.DefaultFontName, "Calibri", "Arial", "Liberation Sans", "DejaVu Sans" })
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

        // ---------------------------------------------------------------------------------------------- drawing

        /// <summary>
        /// Draws a block of rows: gridlines when printed, fills, text, then borders on top.
        /// </summary>
        private void DrawRows(PdfPage page, int r1, int r2, int c1, int c2, double x0, double y0)
        {
            var sheet = _sheet;
            var s = _scale;
            var xs = new double[c2 - c1 + 2];
            xs[0] = x0;
            for (var c = c1; c <= c2; c++) { xs[c - c1 + 1] = xs[c - c1] + (ColumnWidth(c) * s); }
            var ys = new double[r2 - r1 + 2];
            ys[0] = y0;
            for (var r = r1; r <= r2; r++) { ys[r - r1 + 1] = ys[r - r1] + (RowHeight(r) * s); }

            if (sheet.PrintGridLines)
            {
                for (var i = 0; i < xs.Length; i++) { page.DrawLine(xs[i], ys[0], xs[i], ys[ys.Length - 1], "#D9D9D9", 0.25); }
                for (var i = 0; i < ys.Length; i++) { page.DrawLine(xs[0], ys[i], xs[xs.Length - 1], ys[i], "#D9D9D9", 0.25); }
            }

            for (var r = r1; r <= r2; r++)
            {
                if (RowHeight(r) <= 0) { continue; }
                for (var c = c1; c <= c2; c++)
                {
                    var key = XlsxSheet.Key(r, c);
                    if (_covered.Contains(key) || ColumnWidth(c) <= 0 || !sheet.Cells.TryGetValue(key, out var cell)) { continue; }

                    // A merged area spans its columns and rows on this page.
                    double x = xs[c - c1], y = ys[r - r1], w = xs[c - c1 + 1] - x, h = ys[r - r1 + 1] - y;
                    if (_merged.TryGetValue(key, out var end))
                    {
                        w = xs[Math.Min(end.Col2, c2) - c1 + 1] - x;
                        h = ys[Math.Min(end.Row2, r2) - r1 + 1] - y;
                    }

                    Cell(page, cell, r, c, x, y, w, h, c2 - c, xs, c1);
                }
            }
        }

        private void Cell(PdfPage page, XlsxSheet.Cell cell, int row, int column, double x, double y, double w, double h, int room, double[] xs, int c1)
        {
            var style = cell.Style;
            if (style.Fill != null) { page.FillRectangle(x, y, w, h, style.Fill); }

            var text = XlsxNumberText.Format(cell.Value, style.NumberFormat);
            if (text.Length > 0)
            {
                var font = Font(style.FontName, style.Bold, style.Italic);
                var size = (style.FontSize > 0 ? style.FontSize : _book.DefaultFontSize) * _scale;
                var color = style.FontColor ?? "#000000";
                var number = cell.Value is double;
                var horizontal = style.Horizontal ?? (number ? "right" : cell.Value is bool ? "center" : "left");
                var indent = style.Indent * 7.5 * _scale;
                var inner = w - (2 * PadX * _scale) - indent;

                // Lines: wrapped to the cell, or one line that may run into empty cells on the right.
                var lines = style.Wrap ? Wrap(text, font, size, inner) : new List<string> { text };
                var width = inner;
                if (!style.Wrap && !number && horizontal == "left")
                {
                    for (var c = column + 1; c - column <= room && font.Measure(text, size) > width; c++)
                    {
                        if (_sheet.Cells.ContainsKey(XlsxSheet.Key(row, c)) || _covered.Contains(XlsxSheet.Key(row, c))) { break; }
                        width += xs[c - c1 + 1] - xs[c - c1];
                    }
                }

                if (number && !style.Wrap && font.Measure(text, size) > inner)
                {
                    // A number too wide for its column shows hashes, as in Excel.
                    var hashes = Math.Max(1, (int)(inner / Math.Max(0.1, font.Measure("#", size))));
                    lines = new List<string> { new string('#', hashes) };
                }

                var lineHeight = font.LineHeight(size);
                var block = lineHeight * lines.Count;
                var vertical = style.Vertical ?? "bottom";
                var top = vertical == "top" ? y + 1 : vertical == "center" ? y + ((h - block) / 2) : y + h - block - 1;
                page.PushClip(x, y, Math.Max(w, width + (2 * PadX * _scale) + indent), h);
                for (var i = 0; i < lines.Count; i++)
                {
                    var baseline = top + (lineHeight * i) + font.Ascent(size);
                    var line = lines[i];
                    double at;
                    PdfAlign align;
                    if (horizontal == "center" || horizontal == "centerContinuous") { at = x + (w / 2); align = PdfAlign.Center; }
                    else if (horizontal == "right") { at = x + w - (PadX * _scale) - indent; align = PdfAlign.Right; }
                    else { at = x + (PadX * _scale) + indent; align = PdfAlign.Left; }
                    if (cell.Value is XlsxRichText rich && lines.Count == 1)
                    {
                        Rich(page, rich, at, baseline, size, color, align);
                        continue;
                    }

                    page.DrawText(line, at, baseline, font, size, color, align);
                    if (style.Underline || style.DoubleUnderline || style.Strikethrough)
                    {
                        var length = font.Measure(line, size);
                        var start = align == PdfAlign.Left ? at : align == PdfAlign.Center ? at - (length / 2) : at - length;
                        var thick = Math.Max(0.4, size * 0.06);
                        if (style.Underline || style.DoubleUnderline) { page.DrawLine(start, baseline + (size * 0.12), start + length, baseline + (size * 0.12), color, thick); }
                        if (style.DoubleUnderline) { page.DrawLine(start, baseline + (size * 0.22), start + length, baseline + (size * 0.22), color, thick); }
                        if (style.Strikethrough) { page.DrawLine(start, baseline - (size * 0.28), start + length, baseline - (size * 0.28), color, thick); }
                    }
                }

                page.PopClip();
            }

            Border(page, style.Top, x, y, x + w, y);
            Border(page, style.Bottom, x, y + h, x + w, y + h);
            Border(page, style.Left, x, y, x, y + h);
            Border(page, style.Right, x + w, y, x + w, y + h);

            foreach (var (r, c, url) in _sheet.UrlLinks)
            {
                if (r == row && c == column) { page.AddLink(x, y, w, h, url); }
            }
        }

        /// <summary>
        /// Rich text in one line: each run in its own font, size and colour.
        /// </summary>
        private void Rich(PdfPage page, XlsxRichText rich, double at, double baseline, double size, string color, PdfAlign align)
        {
            var parts = new List<(string Text, PdfFont Font, double Size, string Color)>();
            var total = 0.0;
            foreach (var run in rich.Runs)
            {
                var look = run.Look ?? default;
                var font = Font(look.FontName, look.Bold, look.Italic);
                var runSize = look.FontSize > 0 ? look.FontSize * _scale : size;
                parts.Add((run.Text, font, runSize, look.FontColor ?? color));
                total += font.Measure(run.Text, runSize);
            }

            var x = align == PdfAlign.Left ? at : align == PdfAlign.Center ? at - (total / 2) : at - total;
            foreach (var (text, font, runSize, runColor) in parts)
            {
                page.DrawText(text, x, baseline, font, runSize, runColor);
                x += font.Measure(text, runSize);
            }
        }

        private static List<string> Wrap(string text, PdfFont font, double size, double width)
        {
            var lines = new List<string>();
            foreach (var paragraph in text.Replace("\r", string.Empty).Split('\n'))
            {
                var line = string.Empty;
                foreach (var word in paragraph.Split(' '))
                {
                    var candidate = line.Length == 0 ? word : line + " " + word;
                    if (line.Length > 0 && font.Measure(candidate, size) > width)
                    {
                        lines.Add(line);
                        line = word;
                    }
                    else
                    {
                        line = candidate;
                    }
                }

                lines.Add(line);
            }

            return lines;
        }

        /// <summary>
        /// One border line in Excel's styles: thin, medium, thick, dashed, dotted, double (two thin lines).
        /// </summary>
        private void Border(PdfPage page, XlsxBorder border, double x1, double y1, double x2, double y2)
        {
            if (border.Line == null || border.Line == "none")
            {
                return;
            }

            var color = border.Color ?? "#000000";
            var line = border.Line;
            var width = (line.StartsWith("medium", StringComparison.Ordinal) ? 1 : line == "thick" ? 1.5 : line == "hair" ? 0.25 : 0.5) * Math.Max(0.5, _scale);
            var dash = line.IndexOf("ash", StringComparison.Ordinal) >= 0 ? PdfDash.Dashed : line == "dotted" || line == "hair" ? PdfDash.Dotted : PdfDash.Solid;
            if (line == "double")
            {
                var dx = y1 == y2 ? 0 : 0.75;
                var dy = y1 == y2 ? 0.75 : 0;
                page.DrawLine(x1 - dx, y1 - dy, x2 - dx, y2 - dy, color, 0.4).DrawLine(x1 + dx, y1 + dy, x2 + dx, y2 + dy, color, 0.4);
                return;
            }

            page.DrawLine(x1, y1, x2, y2, color, width, dash);
        }

        /// <summary>
        /// Charts and images whose top left corner is on the page, at their anchors.
        /// </summary>
        private void DrawObjects(PdfPage page, int r1, int r2, int c1, int c2, double x0, double y0)
        {
            double X(int column)
            {
                var x = x0;
                for (var c = c1; c < column; c++) { x += ColumnWidth(c) * _scale; }
                return x;
            }

            double Y(int row)
            {
                var y = y0;
                for (var r = r1; r < row; r++) { y += RowHeight(r) * _scale; }
                return y;
            }

            foreach (var chart in _sheet.Charts)
            {
                if (chart.Row >= r1 && chart.Row <= r2 && chart.Column >= c1 && chart.Column <= c2)
                {
                    page.DrawChart(chart, X(chart.Column), Y(chart.Row), chart.Width * 0.75 * _scale, chart.Height * 0.75 * _scale, _book.DefaultFontName);
                }
            }

            foreach (var image in _sheet.Images)
            {
                if (image.Row < r1 || image.Row > r2 || image.Column < c1 || image.Column > c2)
                {
                    continue;
                }

                double x = X(image.Column), y = Y(image.Row), w, h;
                if (image.LastRow > 0 && image.LastColumn > 0)
                {
                    w = X(Math.Min(image.LastColumn, c2) + 1) - x;
                    h = Y(Math.Min(image.LastRow, r2) + 1) - y;
                }
                else
                {
                    w = image.Width * 0.75 * _scale;
                    h = image.Height * 0.75 * _scale;
                }

                try
                {
                    page.DrawImage(image.Data, x, y, w, h);
                }
                catch (NotSupportedException)
                {
                    // A GIF cannot go into a PDF yet: its place stays empty.
                }
            }
        }

        /// <summary>
        /// The sheet's header and footer, with Excel's codes: left, centre and right parts, page, pages, date, time,
        /// sheet and file name; formatting codes are left out.
        /// </summary>
        private void HeaderFooter(PdfPage page, int number, int total, double paperWidth, double paperHeight)
        {
            var sheet = _sheet;
            var m = sheet.Margins;
            var header = number == 1 && sheet.FirstPageHeader != null ? sheet.FirstPageHeader : number % 2 == 0 && sheet.EvenPageHeader != null ? sheet.EvenPageHeader : sheet.Header;
            var footer = number == 1 && sheet.FirstPageFooter != null ? sheet.FirstPageFooter : number % 2 == 0 && sheet.EvenPageFooter != null ? sheet.EvenPageFooter : sheet.Footer;
            var font = Font(null, false, false);
            var size = _book.DefaultFontSize;
            void Parts(string? codes, double baseline)
            {
                if (string.IsNullOrEmpty(codes)) { return; }
                foreach (var (part, text) in Sections(codes!, number, total))
                {
                    var x = part == 'L' ? m.Left * 72 : part == 'R' ? paperWidth - (m.Right * 72) : paperWidth / 2;
                    page.DrawText(text, x, baseline, font, size, "#000000", part == 'L' ? PdfAlign.Left : part == 'R' ? PdfAlign.Right : PdfAlign.Center);
                }
            }

            Parts(header, (m.Header * 72) + font.Ascent(size));
            Parts(footer, paperHeight - (m.Footer * 72) - font.Descent(size));
        }

        private IEnumerable<(char Part, string Text)> Sections(string codes, int number, int total)
        {
            var part = 'C';
            var text = new System.Text.StringBuilder();
            var result = new List<(char, string)>();
            for (var i = 0; i < codes.Length; i++)
            {
                var c = codes[i];
                if (c != '&' || i + 1 >= codes.Length)
                {
                    text.Append(c);
                    continue;
                }

                var code = codes[++i];
                switch (char.ToUpperInvariant(code))
                {
                    case 'L':
                    case 'C':
                    case 'R':
                        if (text.Length > 0) { result.Add((part, text.ToString())); text.Clear(); }
                        part = char.ToUpperInvariant(code);
                        break;
                    case 'P': text.Append(number.ToString(CultureInfo.InvariantCulture)); break;
                    case 'N': text.Append(total.ToString(CultureInfo.InvariantCulture)); break;
                    case 'D': text.Append((_book.Created ?? DateTime.UtcNow).ToString("dd.MM.yyyy", CultureInfo.InvariantCulture)); break;
                    case 'T': text.Append((_book.Created ?? DateTime.UtcNow).ToString("HH:mm", CultureInfo.InvariantCulture)); break;
                    case 'A': text.Append(_sheet.Name); break;
                    case 'F': text.Append(_book.Title ?? "Book"); break;
                    case '&': text.Append('&'); break;
                    case '"':
                        var end = codes.IndexOf('"', i + 1);
                        i = end < 0 ? codes.Length : end;
                        break;
                    default:
                        // Sizes (&12) and switches (&B, &I, &U) change only the look.
                        while (i + 1 < codes.Length && char.IsDigit(codes[i + 1])) { i++; }
                        break;
                }
            }

            if (text.Length > 0) { result.Add((part, text.ToString())); }
            return result;
        }
    }
}
