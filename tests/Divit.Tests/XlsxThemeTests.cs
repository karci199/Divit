namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

/// <summary>
/// The theme part and Excel's built-in cell styles.
/// </summary>
public class XlsxThemeTests
{
    private static string Save(XlsxWorkbook book, string name, bool keep = false)
    {
        book.Created ??= XlsxTestKit.FixedTime;
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    [Fact]
    public void EveryWorkbookHasATheme()
    {
        var book = new XlsxWorkbook { DefaultFontName = "Arial" };
        book.AddSheet("One").Set(1, 1, 1.0);
        var path = Save(book, "theme-default-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var theme = document.WorkbookPart!.ThemePart!.Theme;
            var colors = theme.ThemeElements!.ColorScheme!;
            Assert.Equal("4472C4", colors.Accent1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("70AD47", colors.Accent6Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("0563C1", colors.Hyperlink!.RgbColorModelHex!.Val!.Value);
            var fonts = theme.ThemeElements.FontScheme!;
            Assert.Equal("Arial", fonts.MajorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal("Arial", fonts.MinorFont!.LatinFont!.Typeface!.Value);
        }
    }

    [Fact]
    public void ACustomThemeIsWritten()
    {
        var book = new XlsxWorkbook();
        book.Theme.Name = "Company";
        book.Theme.Accent1 = "#1f4e9a";
        book.Theme.Dark2 = "222222";
        book.Theme.HeadingFont = "Georgia";
        book.Theme.BodyFont = "Verdana";
        var sheet = book.AddSheet("Palette");
        var accents = new[] { book.Theme.Accent1, book.Theme.Accent2, book.Theme.Accent3, book.Theme.Accent4, book.Theme.Accent5, book.Theme.Accent6 };
        for (var i = 0; i < accents.Length; i++)
        {
            sheet.Set(i + 1, 1, "Accent " + (i + 1));
            sheet.Format(i + 1, 2, (ref XlsxStyle s) => s.Fill = accents[i]);
        }

        var path = Save(book, "theme.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var theme = document.WorkbookPart!.ThemePart!.Theme;
            Assert.Equal("Company", theme.Name!.Value);
            Assert.Equal("1F4E9A", theme.ThemeElements!.ColorScheme!.Accent1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("222222", theme.ThemeElements.ColorScheme.Dark2Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("Georgia", theme.ThemeElements.FontScheme!.MajorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal("Verdana", theme.ThemeElements.FontScheme.MinorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal(3, theme.ThemeElements.FormatScheme!.FillStyleList!.Elements<A.SolidFill>().Count());

            var fills = document.WorkbookPart.WorkbookStylesPart!.Stylesheet.Fills!.Elements<Fill>()
                .Select(f => f.PatternFill?.ForegroundColor?.Rgb?.Value).ToList();
            Assert.Contains("FF1F4E9A", fills);
        }
    }

    [Fact]
    public void ABadThemeColourStopsTheSave()
    {
        var book = new XlsxWorkbook();
        book.AddSheet("One");
        book.Theme.Accent3 = "grey";
        var folder = XlsxTestKit.NewFolder();
        Assert.Throws<ArgumentException>(() => book.Save(Path.Combine(folder, "bad.xlsx")));
        Assert.Empty(Directory.GetFiles(folder));
    }

    [Fact]
    public void BuiltInCellStyles()
    {
        var book = new XlsxWorkbook { Title = "Cell styles" };
        var sheet = book.AddSheet("Cell styles");
        var styles = Enum.GetValues(typeof(XlsxCellStyle)).Cast<XlsxCellStyle>().ToArray();
        for (var i = 0; i < styles.Length; i++)
        {
            sheet.Set(i + 1, 1, styles[i].ToString());
            sheet.Set(i + 1, 2, 1234.5);
            sheet.Format(i + 1, 2, (ref XlsxStyle s) => { s.NumberFormat = "#,##0.00"; s.Horizontal = "center"; });
            sheet.ApplyCellStyle(i + 1, 2, i + 1, 2, styles[i]);
        }

        // The same style twice, and a preset given through SetStyle.
        sheet.ApplyCellStyle(1, 3, 2, 3, XlsxCellStyle.Good);
        sheet.SetStyle(3, 3, 3, 3, XlsxStyle.Preset(XlsxCellStyle.Bad));
        sheet.ColumnWidth(1, 14);
        sheet.ColumnWidth(2, 16);
        for (var r = 1; r <= styles.Length; r++)
        {
            sheet.RowHeight(r, 24);
        }

        var path = Save(book, "cell-styles.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            var named = stylesheet.CellStyles!.Elements<CellStyle>().ToList();
            Assert.Equal(styles.Length, named.Count);
            Assert.Equal(("Normal", 0u, 0u), (named[0].Name!.Value!, named[0].FormatId!.Value, named[0].BuiltinId!.Value));
            Assert.Contains(named, n => n.Name == "Heading 1" && n.BuiltinId!.Value == 16);
            Assert.Contains(named, n => n.Name == "Explanatory Text" && n.BuiltinId!.Value == 53);
            Assert.Contains(named, n => n.Name == "Check Cell" && n.BuiltinId!.Value == 23);
            Assert.Equal(named.Count, named.Select(n => n.Name!.Value).Distinct().Count());
            Assert.Equal((uint)named.Count, stylesheet.CellStyleFormats!.Count!.Value);

            var cellFormats = stylesheet.CellFormats!.Elements<CellFormat>().ToList();
            var styleFormats = stylesheet.CellStyleFormats.Elements<CellFormat>().ToList();
            var fills = stylesheet.Fills!.Elements<Fill>().ToList();
            var fonts = stylesheet.Fonts!.Elements<Font>().ToList();
            var cells = document.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!);

            CellFormat FormatOf(string reference) => cellFormats[(int)cells[reference].StyleIndex!.Value];
            CellStyle NamedOf(string reference) => named.Single(n => n.FormatId!.Value == FormatOf(reference).FormatId!.Value);

            // The Good row: the cell points at the Good style, has its look, and keeps its number format and alignment.
            var goodRow = Array.IndexOf(styles, XlsxCellStyle.Good) + 1;
            var good = FormatOf("B" + goodRow);
            Assert.Equal("Good", NamedOf("B" + goodRow).Name!.Value);
            Assert.Equal("FFC6EFCE", fills[(int)good.FillId!.Value].PatternFill!.ForegroundColor!.Rgb!.Value);
            Assert.Equal("FF006100", fonts[(int)good.FontId!.Value].Color!.Rgb!.Value);
            Assert.Equal(HorizontalAlignmentValues.Center, good.Alignment!.Horizontal!.Value);
            Assert.True(good.NumberFormatId!.Value >= 164);

            // The style record itself carries the same look.
            var goodStyle = styleFormats[(int)good.FormatId!.Value];
            Assert.Equal(good.FillId.Value, goodStyle.FillId!.Value);
            Assert.Equal(good.FontId.Value, goodStyle.FontId!.Value);

            // A cell without a style uses record 0, which belongs to Normal.
            Assert.Null(cells["A1"].StyleIndex);
            Assert.Equal(0u, cellFormats[0].FormatId!.Value);
            Assert.Equal("Good", NamedOf("C1").Name!.Value);
            Assert.Equal("Good", NamedOf("C2").Name!.Value);
            Assert.Equal("Bad", NamedOf("C3").Name!.Value);

            var heading = FormatOf("B" + (Array.IndexOf(styles, XlsxCellStyle.Heading1) + 1));
            var headingFont = fonts[(int)heading.FontId!.Value];
            Assert.Equal(15, headingFont.FontSize!.Val!.Value);
            Assert.NotNull(headingFont.Bold);
            var borders = stylesheet.Borders!.Elements<Border>().ToList();
            Assert.Equal(BorderStyleValues.Thick, borders[(int)heading.BorderId!.Value].BottomBorder!.Style!.Value);
        }
    }

    [Fact]
    public void WorkbooksWithoutCellStylesKeepOnlyNormal()
    {
        var book = new XlsxWorkbook();
        book.AddSheet("One").Format(1, 1, (ref XlsxStyle s) => s.Bold = true);
        var path = Save(book, "normal-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var stylesheet = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
            Assert.Equal("Normal", stylesheet.CellStyles!.Elements<CellStyle>().Single().Name!.Value);
            Assert.All(stylesheet.CellFormats!.Elements<CellFormat>(), f => Assert.Equal(0u, f.FormatId!.Value));
        }
    }

    [Fact]
    public void PresetsCarryTheirStyle()
    {
        Assert.Equal(XlsxCellStyle.Total, XlsxStyle.Preset(XlsxCellStyle.Total).CellStyle);
        Assert.NotEqual(XlsxStyle.Preset(XlsxCellStyle.Good), new XlsxStyle { Fill = "#C6EFCE", FontColor = "#006100" });
        Assert.True(XlsxStyle.Preset(XlsxCellStyle.Normal).IsDefault);
    }
}
