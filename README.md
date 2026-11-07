# Divit

A dependency-free .NET library that writes Office documents. It builds the files directly from their XML parts, so
it needs neither Office nor any NuGet package at run time.

*Divit* is the Ottoman scribe's pen case: the pen and the inkwell carried in one piece.

## Status

| Format | State |
|---|---|
| Excel (`.xlsx`) | Writing: text, numbers, dates, booleans, styles (strikethrough, double underline, super/subscript, indent, rotation, shrink to fit), merges, borders, column widths, hidden and grouped rows and columns, frozen panes, auto filter, colour scales, comments, images, links (sheets, files, web, e-mail), defined names, hidden sheets, tab order and active tab, sheet protection, print settings (paper, margins, centring, gridlines), scatter and polar charts |
| Word (`.docx`) | Planned |

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

## Messages

Error messages come from resources: English by default, Turkish (`tr`) as a satellite assembly. They follow the
thread's UI culture.

## Building and testing

```bash
dotnet build Divit.sln
dotnet test Divit.sln
```

Tests run on .NET Framework 4.8 and .NET 10. Every generated file is checked with the Open XML SDK validator, which is
a test-only dependency. Sample files land in `artifacts/gallery` for a look in Excel.

## Roadmap

1. Excel: formulas, conditional formatting rules, tables (ListObject), more chart kinds (column, bar, line, area, pie),
   data validation, rich text, sparklines, streaming for very large sheets
2. Shared package layer (zip, content types, relationships, document properties) for every format
3. Word (`.docx`): paragraphs, styles, tables, images, headers and footers, page setup
4. Later: PowerPoint (`.pptx`), SVG charts, HTML, CSV, PDF
5. Macros last, if ever: a prebuilt `vbaProject.bin` or COM
