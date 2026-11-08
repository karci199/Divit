namespace Divit.Tests;

using Divit.Xlsx;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Xunit;
using Reader = UglyToad.PdfPig.PdfDocument;

/// <summary>
/// Workbooks printed as PDF: number formats as Excel shows them, and a measurement log over several pages with its
/// title rows, styles, merges, chart and footer.
/// </summary>
public class XlsxPdfTests
{
    [Theory]
    [InlineData(1234.5, null, "1234.5")]
    [InlineData(0.1 + 0.2, "General", "0.3")]
    [InlineData(1234.5, "#,##0.00", "1,234.50")]
    [InlineData(-1234.5, "#,##0.00;(#,##0.00)", "(1,234.50)")]
    [InlineData(0.125, "0.0%", "12.5%")]
    [InlineData(0, "0.00;-0.00;\"–\"", "–")]
    [InlineData(12345678, "0.00E+00", "1.23E+07")]
    [InlineData(8.25, "0.00 \"GHz\"", "8.25 GHz")]
    [InlineData(46024.0, "yyyy-mm-dd", "2026-01-02")]
    [InlineData(46024.5625, "dd.mm.yyyy hh:mm", "02.01.2026 13:30")]
    [InlineData(46024.5625, "h:mm AM/PM", "1:30 PM")]
    [InlineData(46024.0, "d mmm yyyy", "2 Jan 2026")]
    [InlineData(1.5, "[Red]0.0", "1.5")]
    [InlineData(2.0, "0.0_)", "2.0 ")]
    [InlineData(-0.00004, "0.0%", "0.0%")]
    public void NumbersShowAsInExcel(double value, string? format, string expected)
    {
        Assert.Equal(expected, XlsxNumberText.Format(value, format));
    }

    [Fact]
    public void OtherValuesShowAsInExcel()
    {
        Assert.Equal("TRUE", XlsxNumberText.Format(true, null));
        Assert.Equal("Text", XlsxNumberText.Format("Text", "0.00"));
        Assert.Equal("Item: Text", XlsxNumberText.Format("Text", "\"Item: \"@"));
        Assert.Equal(string.Empty, XlsxNumberText.Format(null, null));
    }

    /// <summary>
    /// A measurement log: a merged title, a header row repeated on every page, numbers, percentages and dates, a
    /// wrapped remark, borders, a chart, and a footer with page numbers.
    /// </summary>
    internal static XlsxWorkbook Log()
    {
        var book = new XlsxWorkbook { Created = XlsxTestKit.FixedTime, Title = "Measurement log" };
        book.DefaultFontName = "Trebuchet MS";
        var sheet = book.AddSheet("Log");
        sheet.Set(1, 1, "SH2000 return loss log");
        sheet.Merge(1, 1, 1, 6);
        sheet.Format(1, 1, (ref XlsxStyle s) => { s.Bold = true; s.FontSize = 14; s.FontColor = "#1F4E79"; });
        var heads = new[] { "No", "Date", "Frequency", "VSWR", "Margin", "Remark" };
        for (var c = 0; c < heads.Length; c++)
        {
            sheet.Set(2, c + 1, heads[c]);
        }

        sheet.Format(2, 1, 2, 6, (ref XlsxStyle s) =>
        {
            s.Bold = true;
            s.Fill = "#BDD7EE";
            s.FontColor = "#1F3864";
            s.Bottom = new XlsxBorder("medium", "#1F4E79");
            s.Horizontal = "center";
        });
        sheet.ColumnWidth(1, 6);
        sheet.ColumnWidth(2, 12);
        sheet.ColumnWidth(3, 12);
        sheet.ColumnWidth(4, 9);
        sheet.ColumnWidth(5, 9);
        sheet.ColumnWidth(6, 30);
        for (var i = 0; i < 120; i++)
        {
            var row = i + 3;
            var f = 2 + (i * 0.25);
            var v = 1.30 + (0.10 * Math.Sin(1.9 * f)) + (0.75 * Math.Exp(-Math.Pow((f - 29.6) / 1.1, 2)));
            sheet.Set(row, 1, i + 1);
            sheet.Set(row, 2, new DateTime(2026, 1, 2).AddDays(i / 40));
            sheet.Set(row, 3, f);
            sheet.Set(row, 4, v);
            sheet.Set(row, 5, (2 - v) / 2);
            sheet.Format(row, 2, (ref XlsxStyle s) => s.NumberFormat = "dd.mm.yyyy");
            sheet.Format(row, 3, (ref XlsxStyle s) => s.NumberFormat = "0.00 \"GHz\"");
            sheet.Format(row, 4, (ref XlsxStyle s) => { s.NumberFormat = "0.00"; s.FontColor = v > 2 ? "#C00000" : null; s.Bold = v > 2; });
            sheet.Format(row, 5, (ref XlsxStyle s) => s.NumberFormat = "0.0%");
            if (v > 2)
            {
                sheet.Set(row, 6, "Above the limit of 2.00: repeat after reseating the connector and recalibrating port 1.");
                sheet.Format(row, 6, (ref XlsxStyle s) => s.Wrap = true);
                sheet.RowHeight(row, 42);
            }
        }

        sheet.Frame(2, 1, 122, 6, new XlsxBorder("thin", "#9DC3E6"), new XlsxBorder("thin", "#DDEBF7"));
        sheet.PrintTitleRows = (2, 2);
        sheet.Footer = "&LSH2000 log&CPage &P of &N&R&A";
        var chart = sheet.AddChart(125, 1, 620, 300);
        chart.Title = "VSWR over frequency";
        chart.AddSeries("VSWR", 3, 122, 3, 4);
        chart.XAxis.Title = "Frequency (GHz)";
        return book;
    }

    [Fact]
    public void ALogPrintsOverPages()
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), "measurement-log.pdf");
        Log().SavePdf(path);
        Assert.Empty(PdfTestKit.Validate(path));
        XlsxTestKit.Keep(path, "measurement-log.pdf");

        using (var reader = Reader.Open(path))
        {
            var pages = reader.GetPages().ToList();
            Assert.True(pages.Count >= 3, pages.Count + " pages");
            for (var i = 0; i < pages.Count; i++)
            {
                var text = pages[i].Text;
                Assert.Contains("Page " + (i + 1).ToString(CultureInfo.InvariantCulture) + " of " + pages.Count.ToString(CultureInfo.InvariantCulture), text);
                if (i < pages.Count - 1)
                {
                    Assert.Contains("FrequencyVSWR", text.Replace(" ", string.Empty));
                }
            }

            var all = string.Concat(pages.Select(p => p.Text));
            Assert.Contains("SH2000 return loss log", all);
            Assert.Contains("02.01.2026", all);
            Assert.Contains("2.00 GHz", all);
            Assert.Contains("%", all);
            Assert.Contains("Above the limit", all);
            Assert.Contains("VSWR over frequency", all);
            for (var n = 1; n <= 120; n++)
            {
                Assert.Contains(pages.SelectMany(p => p.GetWords()), w => w.Text == n.ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    [Fact]
    public void WidePagesFitTheirWidth()
    {
        var book = new XlsxWorkbook { Created = XlsxTestKit.FixedTime };
        var sheet = book.AddSheet("Wide");
        for (var c = 1; c <= 30; c++)
        {
            sheet.Set(1, c, "Column " + c.ToString(CultureInfo.InvariantCulture));
            sheet.ColumnWidth(c, 15);
        }

        var path = Path.Combine(XlsxTestKit.NewFolder(), "wide.pdf");
        book.SavePdf(path);
        Assert.Empty(PdfTestKit.Validate(path));
        using (var reader = Reader.Open(path))
        {
            Assert.Equal(1, reader.NumberOfPages);
            var words = reader.GetPage(1).GetWords().ToList();
            Assert.Contains(words, w => w.Text == "30");
            Assert.True(words.Max(w => w.BoundingBox.Right) <= 595.28 - 35, "fits between the margins");
        }

        sheet.FitWidth = 0;
        book.SavePdf(path);
        using (var reader = Reader.Open(path))
        {
            Assert.True(reader.NumberOfPages > 1, "without fitting the columns go over several pages");
        }
    }

    [Fact]
    public void PdfsOfWorkbooksAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        Log().SavePdf(Path.Combine(folder, "a.pdf"));
        Log().SavePdf(Path.Combine(folder, "b.pdf"));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.pdf")), File.ReadAllBytes(Path.Combine(folder, "b.pdf")));
    }
}
