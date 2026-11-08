namespace Divit.Tests;

using Divit.Charts;
using Divit.Docx;
using Divit.Pdf;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Xunit;
using Reader = UglyToad.PdfPig.PdfDocument;

/// <summary>
/// Charts from values: drawn on PDF pages, written as real charts into Word documents with their data embedded,
/// and drawn again when those documents are saved as PDF.
/// </summary>
public class ChartTests
{
    /// <summary>
    /// The VSWR of the horn every 50 MHz from 2 to 32 GHz, with its limit and its highest point.
    /// </summary>
    internal static Chart Vswr()
    {
        var f = new List<double>();
        var v = new List<double>();
        for (var i = 0; i <= 600; i++)
        {
            var x = 2 + (i * 0.05);
            f.Add(x);
            v.Add(1.30 + (0.10 * Math.Sin(1.9 * x)) + (0.06 * Math.Sin(0.7 * x)) + (0.75 * Math.Exp(-Math.Pow((x - 29.6) / 1.1, 2))));
        }

        var peak = v.IndexOf(v.Max());
        var chart = new Chart(ChartKind.Scatter) { Title = "VSWR over frequency" };
        chart.XAxis.Title = "Frequency (GHz)";
        chart.XAxis.Min = 2;
        chart.XAxis.Max = 32;
        chart.XAxis.MajorUnit = 2;
        chart.YAxis.Title = "VSWR";
        chart.YAxis.Min = 1;
        chart.YAxis.Max = 2.5;
        chart.YAxis.NumberFormat = "0.00";
        chart.AddSeries("VSWR", f, v).Width = 1.5;
        chart.AddLimit("Upper limit 2.00", 2.0);
        chart.AddPoints("Maximum", new[] { f[peak] }, new[] { v[peak] });
        return chart;
    }

    private static Chart Gains()
    {
        var bands = new[] { "2–8 GHz", "8–18 GHz", "18–26 GHz", "26–32 GHz" };
        var chart = new Chart(ChartKind.Column) { Title = "Gain per band", Subtitle = "Three antennas" };
        chart.YAxis.Title = "Gain (dBi)";
        chart.AddSeries("SH2000-01", bands, new[] { 9.1, 12.4, 14.8, 16.2 });
        chart.AddSeries("SH2000-02", bands, new[] { 8.8, 12.1, 15.0, 16.5 });
        chart.AddSeries("SH2000-03", bands, new[] { 9.3, 12.6, 14.4, 15.9 });
        chart.AddLimit("Minimum", 10);
        return chart;
    }

    private static Chart Results()
    {
        var chart = new Chart(ChartKind.Doughnut) { Title = "Test results", ShowLegend = true };
        chart.AddSeries("Steps", new[] { "Passed", "Failed", "Not run" }, new[] { 41.0, 3, 4 });
        return chart;
    }

    private static Chart Stacked()
    {
        var months = new[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        var chart = new Chart(ChartKind.Bar) { Title = "Antennas tested per month" };
        chart.AddSeries("Accepted", months, new[] { 12.0, 15, 9, 18, 20, 17 });
        chart.AddSeries("Rejected", months, new[] { 1.0, 2, 0, 3, 1, 2 });
        return chart;
    }

    private static Chart Trend()
    {
        var weeks = Enumerable.Range(1, 12).Select(w => "W" + w.ToString(CultureInfo.InvariantCulture)).ToArray();
        var chart = new Chart(ChartKind.Area) { Title = "Chamber hours per week" };
        chart.AddSeries("Hours", weeks, weeks.Select((_, i) => 20 + (10 * Math.Sin(i * 0.6)) + i).ToArray());
        return chart;
    }

    [Fact]
    public void ChartsAreDrawnOnPdfPages()
    {
        var doc = new PdfDocument { Created = XlsxTestKit.FixedTime, Title = "Charts" };
        var page = doc.AddPage();
        page.DrawChart(Vswr(), 40, 40, 515, 300, "Trebuchet MS");
        page.DrawChart(Gains(), 40, 360, 300, 220, "Trebuchet MS");
        page.DrawChart(Results(), 350, 360, 205, 220, "Trebuchet MS");
        page.DrawChart(Stacked(), 40, 600, 250, 200, "Trebuchet MS");
        page.DrawChart(Trend(), 305, 600, 250, 200, "Trebuchet MS");
        var path = Path.Combine(XlsxTestKit.NewFolder(), "charts.pdf");
        doc.Save(path);
        Assert.Empty(PdfTestKit.Validate(path));
        XlsxTestKit.Keep(path, "charts.pdf");

        using (var reader = Reader.Open(path))
        {
            var text = reader.GetPage(1).Text;
            foreach (var expected in new[] { "VSWR over frequency", "Frequency (GHz)", "2.00", "2.50", "Upper limit 2.00", "Gain per band", "18–26 GHz", "SH2000-02", "Passed", "Jun", "W1" })
            {
                Assert.Contains(expected, text);
            }
        }

        // The curve is one line of hundreds of points, kept inside the plot.
        var content = PdfTestKit.Content(page);
        Assert.Contains("W n", content);
        Assert.True(content.Split('\n').Any(l => l.Split(new[] { " l " }, StringSplitOptions.None).Length > 500), "the VSWR curve is drawn");
    }

    [Fact]
    public void WordChartsAreRealChartsWithTheirData()
    {
        var doc = new DocxDocument { Created = XlsxTestKit.FixedTime, Title = "Charts", NumberHeadings = true };
        doc.Look.BodyFont = "Trebuchet MS";
        doc.AddHeading("Return loss");
        doc.AddParagraph().AddChart(Vswr(), 16, 9);
        doc.AddCaption(DocxCaptionKind.Figure, "VSWR of SN 24-0187");
        doc.AddHeading("Gain");
        doc.AddParagraph().AddChart(Gains(), 16, 8);
        doc.AddParagraph().AddChart(Results(), 8, 7).AddChart(Stacked(), 8, 7);
        var path = Path.Combine(XlsxTestKit.NewFolder(), "word-charts.docx");
        doc.Save(path);
        Assert.Empty(DocxTestKit.Validate(path));
        XlsxTestKit.Keep(path, "word-charts.docx");

        using (var document = WordprocessingDocument.Open(path, false))
        {
            var charts = document.MainDocumentPart!.ChartParts.ToList();
            Assert.Equal(4, charts.Count);
            foreach (var chart in charts)
            {
                Assert.NotNull(chart.EmbeddedPackagePart);
                Assert.Contains("externalData", chart.ChartSpace!.OuterXml);
                using (var stream = chart.EmbeddedPackagePart!.GetStream())
                using (var copy = new MemoryStream())
                {
                    stream.CopyTo(copy);
                    copy.Position = 0;
                    using (var book = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(copy, false))
                    {
                        Assert.Single(book.WorkbookPart!.WorksheetParts);
                    }
                }
            }

            var vswr = charts.Single(c => c.ChartSpace!.InnerText.Contains("VSWR over frequency"));
            Assert.Contains("<c:scatterChart>", vswr.ChartSpace!.OuterXml);
            Assert.Equal(601 + 2 + 1, vswr.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.NumericPoint>().Count(p => p.Parent!.Parent!.Parent is DocumentFormat.OpenXml.Drawing.Charts.YValues));
        }

        var pdf = Path.Combine(XlsxTestKit.NewFolder(), "word-charts.pdf");
        doc.SavePdf(pdf);
        Assert.Empty(PdfTestKit.Validate(pdf));
        XlsxTestKit.Keep(pdf, "word-charts.pdf");
        using (var reader = Reader.Open(pdf))
        {
            var text = string.Concat(reader.GetPages().Select(p => p.Text));
            Assert.Contains("VSWR over frequency", text);
            Assert.Contains("Gain per band", text);
            Assert.Contains("Figure 1", text);
        }
    }

    [Fact]
    public void ChartsAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        foreach (var name in new[] { "a", "b" })
        {
            var doc = new DocxDocument { Created = XlsxTestKit.FixedTime };
            doc.AddParagraph().AddChart(Gains());
            doc.Save(Path.Combine(folder, name + ".docx"));
            doc.SavePdf(Path.Combine(folder, name + ".pdf"));
        }

        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.docx")), File.ReadAllBytes(Path.Combine(folder, "b.docx")));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.pdf")), File.ReadAllBytes(Path.Combine(folder, "b.pdf")));
    }
}
