# Divit

A dependency-free .NET library that writes Office documents. It builds the files directly from their XML parts, so
it needs neither Office nor any NuGet package at run time.

*Divit* is the Ottoman scribe's pen case: the pen and the inkwell carried in one piece.

## Status

| Format | State |
|---|---|
| Excel (`.xlsx`) | Writing: text, numbers, dates, booleans, styles (strikethrough, double underline, super/subscript, indent, rotation, shrink to fit), merges, borders, column widths, hidden and grouped rows and columns, frozen panes, auto filter, formulas, data validation, conditional formatting (rules, colour scales, data bars, icon sets), rich text, tables with totals rows, pivot tables (row, column, filter and value fields), theme, built-in cell styles, comments (rich text, box size), images (PNG, JPEG, GIF; at a cell or filling cells), links (sheets, files, web, e-mail), defined names, hidden sheets, tab order and active tab, sheet and workbook protection (SHA-512, editable cells), print settings (paper, margins, centring, gridlines, page breaks, first and even page headers), charts (scatter, polar, column, bar, line, area, pie, doughnut; combined, secondary axis, stacked and 100 %, logarithmic and reversed axes, trendlines, error bars, data labels), streamed sheets for millions of rows |
| Word (`.docx`) | Writing: sections (page size, orientation, margins, vertical alignment, page numbering), paragraphs and fonts, numbered headings (1, 1.1, 1.1.1), a house style turned into Word styles, headers and footers (first page, even pages) with page fields, tables (house, grid and plain looks, repeated header rows, column widths and alignment, merged cells, also in headers and footers), images (PNG, JPEG, GIF), figure and table captions, table of contents and lists of figures and tables, bulleted and numbered lists (nine levels), links (web, e-mail, within the document), cross-references (text, heading number, page, caption label), watermarks, check boxes, footnotes, document protection (read only, forms, comments, tracked changes; SHA-512), change tracking |

Divit writes documents; it does not read or edit existing ones. Build the document again from your own data instead.

## Targets

- `netstandard2.0`: .NET Framework 4.8 and every newer runtime
- `net10.0`: annotated for nullable reference types, trimming and AOT

## Example

```csharp
using Divit.Xlsx;

var book = new XlsxWorkbook { Title = "Measurement" };
var sheet = book.AddSheet("Data");
sheet.Set(1, 1, "Angle");
sheet.Set(1, 2, "Gain (dB)");
for (var i = 0; i < 37; i++)
{
    sheet.Set(i + 2, 1, i * 10.0);
    sheet.Set(i + 2, 2, 10 * Math.Cos(i * 10.0 * Math.PI / 180));
}

var chart = sheet.AddChart(2, 4, 640, 360);
chart.Title = "Gain";
chart.AddSeries("Gain", 2, 38, 1, 2);
book.Save("measurement.xlsx");
```

## Word example

```csharp
using Divit.Docx;

var doc = new DocxDocument { Title = "Test procedure", NumberHeadings = true };
doc.Look.BodyFont = "Trebuchet MS";
doc.Look.BodySize = 12;
doc.Section.Header.AddParagraph().Add("TP-001 Rev. A  ·  Page ").AddField(DocxField.Page);
doc.AddTableOfContents(3, "Contents");
doc.AddHeading("Tests", 1);
doc.AddHeading("Hardware list", 2);
doc.AddCaption(DocxCaptionKind.Table, "Hardware list");
var table = doc.AddTable(3).SetColumnWidths(1, 9, 6);
table.AddHeaderRow("No", "Equipment", "Model");
table.AddRow("1", "Vector network analyser", "ZNA43");
doc.Save("procedure.docx");
```

A full acceptance test procedure (cover with letterhead, headers and footers as tables, change record, contents,
lists of figures and tables, numbered chapters, hardware and software lists, figures, record form) is built in
`tests/Divit.Tests/DocxProcedureTests.cs` from these blocks only.

## Very large sheets

`AddStreamedSheet` writes the rows of a sheet while the file is saved, without keeping them: memory stays flat. Ten
million cells take about 30 MB instead of 2.6 GB, and save about twice as fast.

```csharp
book.AddStreamedSheet("Log", rows =>
{
    rows.Cell("Time").Cell("Value").EndRow();
    foreach (var m in measurements)
    {
        rows.Cell(m.Time).Cell(m.Value).EndRow();
    }
});
```

## Messages

Error messages come from resources: English by default, Turkish (`tr`) as a satellite assembly. They follow the
thread's UI culture.

## Building and testing

```bash
dotnet build Divit.sln
dotnet test Divit.sln
```

Tests run on .NET Framework 4.8 and .NET 10. Every generated file is checked with the Open XML SDK validator, which is
a test-only dependency. Sample files land in `artifacts/gallery` for a look in Excel and Word.

## Roadmap

1. Word (`.docx`): streamed writing for very long documents, then the backlog below
2. Later: PowerPoint (`.pptx`), SVG charts, HTML, CSV, PDF
3. Macros last, if ever: a prebuilt `vbaProject.bin` or COM

### Word backlog

Left out until a use needs them:

- **Floating images and text boxes**: pictures that text flows around, placed anywhere on the page.
- **Native charts**: Word charts built on the Excel chart writer, with their data in an embedded workbook.
- **Tracked insertions and deletions**: text written as a proposed change (`w:ins`, `w:del`) with author and date,
  for example to show what changed between two revisions. Turning tracking on for readers is done.
- **Comments**: notes in the margin on a piece of text.
- **Endnotes**, and footnotes in other styles (letters, symbols, restarting per page).
- **More content controls**: drop-down lists, date pickers, plain text fields for forms.
- **Editable ranges in a protected document**: parts anyone may change while the rest stays locked.
- **Columns**: newspaper-style text in two or three columns.
- **Bibliography and citations**: Word's source manager.
- **Equations**: Office Math.

### Excel backlog

Left out until a use needs them:

- **Sparklines**: a small chart inside a cell. An Excel 2010 extension that some readers ignore.
- **More pivot table layouts**: several row or column fields, several value fields, grouping dates.
- **Slicers and timelines**: filter buttons for tables and pivot tables.
- **Array formulas and dynamic arrays** (`FILTER`, `SORT`, spilling results).
- **Chart sheets**: a sheet that is one chart.
- **Cached values in charts on streamed sheets**: today Excel reads the cells when it opens the file.
- **Smaller files on .NET 10**: its fastest deflate level packs about a third larger than .NET Framework's; a
  stronger level costs save time.
- **Same number text on both runtimes**: .NET Framework's round-trip format writes some fractions with more digits
  than .NET does (same value).
- **Auto-fit column widths**: measuring text needs font metrics.
- **Computing formulas**: Excel computes them on opening; only needed for readers that never do.
