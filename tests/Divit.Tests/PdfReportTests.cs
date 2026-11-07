namespace Divit.Tests;

using Divit.Pdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Xunit;
using Reader = UglyToad.PdfPig.PdfDocument;

/// <summary>
/// A demo built with the drawing API only: the acceptance test report of the SH2000 horn, with its header, test
/// information, overall result, limits table and VSWR chart. The library stays generic; the report is a test.
/// </summary>
public class PdfReportTests
{
    private const string Blue = "#1F4E79", Light = "#BDD7EE", Band = "#EEF5FB", Line = "#9DC3E6", Ink = "#262626", Grey = "#595959";
    private const string Green = "#375623", GreenFill = "#E2F0D9", Red = "#C00000", RedFill = "#FBE5E5";

    private static string F(double value, string format) => value.ToString(format, CultureInfo.InvariantCulture).Replace("-", "−");

    /// <summary>
    /// The measurement: VSWR every 50 MHz from 2 to 32 GHz, with a resonance near 30 GHz.
    /// </summary>
    private static List<(double F, double V)> Vswr()
    {
        var list = new List<(double, double)>();
        for (var i = 0; i <= 600; i++)
        {
            var f = 2 + (i * 0.05);
            list.Add((f, 1.30 + (0.10 * Math.Sin(1.9 * f)) + (0.06 * Math.Sin(0.7 * f)) + (0.75 * Math.Exp(-Math.Pow((f - 29.6) / 1.1, 2)))));
        }

        return list;
    }

    internal static PdfDocument Report()
    {
        var doc = new PdfDocument { Title = "Acceptance Test Report ATR-SH2000-0042", Author = "Divit Electronics", Subject = "SH2000 horn antenna", Created = XlsxTestKit.FixedTime };
        var regular = doc.AddFont("Trebuchet MS");
        var bold = doc.AddFont("Trebuchet MS", bold: true);
        var page = doc.AddPage();
        double left = 40, right = page.Width - 40, wide = right - left;

        var vswr = Vswr();
        var low = vswr.Where(p => p.F <= 18).OrderByDescending(p => p.V).First();
        var high = vswr.Where(p => p.F > 18).OrderByDescending(p => p.V).First();
        var steps = new (string Step, double? Lower, double? Upper, double Value, string Unit, string Format)[]
        {
            ("Return loss at 2 GHz", null, -10, -14.2, "dB", "0.0"),
            ("VSWR, maximum 2–18 GHz", null, 2.0, low.V, "", "0.00"),
            ("VSWR, maximum 18–32 GHz", null, 2.0, high.V, "", "0.00"),
            ("Gain at 10 GHz", 10, 14, 11.8, "dBi", "0.0"),
            ("Gain at 30 GHz", 14, 18, 16.9, "dBi", "0.0"),
            ("3 dB beamwidth, E-plane at 10 GHz", 30, 50, 42.5, "°", "0.0"),
            ("Cross-polar discrimination at 10 GHz", 20, null, 23.4, "dB", "0.0"),
            ("Insulation resistance, connector to body", 100, null, 450, "MΩ", "0"),
        };
        bool Pass((string, double? Lower, double? Upper, double Value, string, string) s) => (s.Lower == null || s.Value >= s.Lower) && (s.Upper == null || s.Value <= s.Upper);
        var passed = steps.Count(Pass);

        // Title band.
        page.FillRectangle(0, 0, page.Width, 92, Blue);
        page.DrawText("Acceptance Test Report", left, 46, bold, 22, "#FFFFFF");
        page.DrawText("SH2000 dual-ridge horn antenna · 2–32 GHz", left, 68, regular, 11, Light);
        page.DrawText("DIVIT ELECTRONICS", right, 44, bold, 12, "#FFFFFF", PdfAlign.Right);
        page.DrawText("Test and Qualification Department", right, 60, regular, 9, Light, PdfAlign.Right);

        // Information in two columns.
        var info = new[]
        {
            ("Report no", "ATR-SH2000-0042"), ("Date", "07.11.2026 14:32"),
            ("Procedure", "TP-SH2000-001 Rev. B"), ("Test station", "Anechoic chamber 2"),
            ("Unit under test", "SH2000 horn antenna"), ("Operator", "Ayşe Yılmaz"),
            ("Serial no", "SN 24-0187"), ("Ambient", "23.1 °C, 41 % RH"),
        };
        for (var i = 0; i < info.Length; i++)
        {
            var x = i % 2 == 0 ? left : left + (wide / 2);
            var y = 122 + ((i / 2) * 18);
            page.DrawText(info[i].Item1, x, y, regular, 9, Grey);
            page.DrawText(info[i].Item2, x + 92, y, bold, 10, Ink);
        }

        // Overall result.
        var ok = passed == steps.Length;
        page.FillRectangle(left, 186, wide, 28, ok ? GreenFill : RedFill).FillRectangle(left, 186, 5, 28, ok ? "#548235" : Red);
        page.DrawText("Overall result: " + (ok ? "PASSED" : "FAILED"), left + 14, 205, bold, 12, ok ? Green : Red);
        page.DrawText(passed + " of " + steps.Length + " steps within limits", right - 10, 205, regular, 10, ok ? Green : Red, PdfAlign.Right);

        // The steps with their limits.
        var columns = new (string Head, double Width, PdfAlign Align)[]
        {
            ("No", 26, PdfAlign.Left), ("Test step", 205, PdfAlign.Left), ("Lower", 52, PdfAlign.Right), ("Upper", 52, PdfAlign.Right),
            ("Measured", 62, PdfAlign.Right), ("Unit", 50, PdfAlign.Center), ("Result", 68, PdfAlign.Center),
        };
        double At(double x, (string, double Width, PdfAlign Align) c) => c.Align == PdfAlign.Left ? x + 6 : c.Align == PdfAlign.Right ? x + c.Width - 6 : x + (c.Width / 2);
        double top = 232, row = 19;
        page.DrawText("Test results", left, top, bold, 13, Blue);
        top += 10;
        page.FillRectangle(left, top, wide, row + 2, Light);
        var cx = left;
        foreach (var c in columns)
        {
            page.DrawText(c.Head, At(cx, c), top + 14, bold, 9, "#1F3864", c.Align);
            cx += c.Width;
        }

        top += row + 2;
        for (var i = 0; i < steps.Length; i++)
        {
            var s = steps[i];
            var pass = Pass(s);
            if (i % 2 == 1)
            {
                page.FillRectangle(left, top, wide, row, Band);
            }

            var cells = new[]
            {
                (i + 1).ToString(CultureInfo.InvariantCulture), s.Step, s.Lower is double lo ? F(lo, s.Format) : "–", s.Upper is double up ? F(up, s.Format) : "–",
                F(s.Value, s.Format), s.Unit.Length == 0 ? "–" : s.Unit, pass ? "PASS" : "FAIL",
            };
            cx = left;
            for (var c = 0; c < columns.Length; c++)
            {
                var last = c == columns.Length - 1;
                if (last)
                {
                    page.FillRectangle(cx + 8, top + 3, columns[c].Width - 16, row - 6, pass ? GreenFill : RedFill);
                }

                var color = last ? (pass ? Green : Red) : c == 4 && !pass ? Red : Ink;
                page.DrawText(cells[c], At(cx, columns[c]), top + 13, c == 4 || last ? bold : regular, 9, color, columns[c].Align);
                cx += columns[c].Width;
            }

            top += row;
            page.DrawLine(left, top, right, top, Line, 0.5);
        }

        // VSWR against frequency with the limit.
        top += 26;
        page.DrawText("VSWR over frequency", left, top, bold, 13, Blue);
        double plotLeft = left + 34, plotRight = right - 8, plotTop = top + 14, plotBottom = top + 262;
        double X(double f) => plotLeft + ((f - 2) / 30 * (plotRight - plotLeft));
        double Y(double v) => plotBottom - ((v - 1.0) / 1.5 * (plotBottom - plotTop));
        for (var v = 1.0; v <= 2.5001; v += 0.25)
        {
            page.DrawLine(plotLeft, Y(v), plotRight, Y(v), "#D9D9D9", 0.4);
            page.DrawText(F(v, "0.00"), plotLeft - 5, Y(v) + 3, regular, 8, Grey, PdfAlign.Right);
        }

        for (var f = 2; f <= 32; f += 2)
        {
            page.DrawLine(X(f), plotTop, X(f), plotBottom, "#E7E6E6", 0.4);
            page.DrawText(f.ToString(CultureInfo.InvariantCulture), X(f), plotBottom + 11, regular, 8, Grey, PdfAlign.Center);
        }

        page.DrawLine(plotLeft, plotBottom, plotRight, plotBottom, Grey, 0.8).DrawLine(plotLeft, plotTop, plotLeft, plotBottom, Grey, 0.8);
        page.DrawLine(X(18), plotTop, X(18), plotBottom, Grey, 0.6, PdfDash.Dotted);
        page.DrawLine(plotLeft, Y(2.0), plotRight, Y(2.0), Red, 1, PdfDash.Dashed);
        page.DrawText("Upper limit 2.00", plotLeft + 6, Y(2.0) - 4, bold, 8, Red);
        page.DrawPolyline(vswr.Select(p => (X(p.F), Y(p.V))).ToList(), "#2E75B6", 1.3);
        page.FillEllipse(X(high.F), Y(high.V), 3, 3, Red);
        page.DrawText("max " + F(high.V, "0.00") + " at " + F(high.F, "0.0") + " GHz", X(high.F) - 6, Y(high.V) - 6, bold, 8, Red, PdfAlign.Right);
        page.DrawText("Frequency (GHz)", (plotLeft + plotRight) / 2, plotBottom + 25, regular, 9, Grey, PdfAlign.Center);

        // Signatures and footer.
        var signs = new[] { ("Tested by", "Ayşe Yılmaz"), ("Checked by", string.Empty), ("Approved by", string.Empty) };
        for (var i = 0; i < signs.Length; i++)
        {
            var x = left + (i * (wide / 3));
            page.DrawText(signs[i].Item1, x, 770, regular, 8, Grey);
            page.DrawText(signs[i].Item2, x, 785, bold, 10, Ink);
            page.DrawLine(x, 792, x + (wide / 3) - 20, 792, Grey, 0.5);
        }

        page.DrawLine(left, 808, right, 808, Line, 0.8);
        page.DrawText("ATR-SH2000-0042 · TP-SH2000-001 Rev. B · generated by Divit", left, 821, regular, 8, Grey);
        page.DrawText("Page 1 of 1", right, 821, regular, 8, Grey, PdfAlign.Right);
        return doc;
    }

    [Fact]
    public void TheReportIsValidAndReadable()
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), "test-report.pdf");
        Report().Save(path);
        Assert.Empty(PdfTestKit.Validate(path));
        XlsxTestKit.Keep(path, "test-report.pdf");

        Assert.True(new FileInfo(path).Length < 60000, "report of " + new FileInfo(path).Length + " bytes");
        using (var reader = Reader.Open(path))
        {
            var text = reader.GetPage(1).Text;
            Assert.Contains("Overall result: FAILED", text);
            Assert.Contains("Ayşe Yılmaz", text);
            Assert.Contains("max 2.09 at 29.8 GHz", text);
            Assert.Contains("−14.2", text);
            Assert.Equal("Acceptance Test Report ATR-SH2000-0042", reader.Information.Title);
        }
    }
}
