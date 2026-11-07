# Divit

Dependency-free .NET library that writes Office documents (xlsx now, docx next). `src/Divit` is the library
(`netstandard2.0;net10.0`), `tests/Divit.Tests` the xUnit tests (`net48;net10.0`). See README.md for the roadmap.
`src/Divit/Packaging` is the internal layer every format shares (zip and atomic save, content types, relationships,
document properties, theme part, image format and size, SHA-512 password hash, allocation-free number formatting);
`src/Divit/Xlsx` builds on it, Word will too. `XlsxWorkbook` is split into partial files by part (`.Package`,
`.Sheet`, `.Styles`, `.Strings`, `.Tables`, `.Pivot`, `.Stream`).

## Rules

1. **Everything is English**: identifiers, comments, XML docs, commit messages, documents. The only exception is
   localised resources: `Properties/Strings.resx` is English (neutral), `Strings.tr.resx` Turkish.
2. **No run-time dependencies.** Only the framework (`System.IO.Compression`, `System.Xml`). Test-only packages are fine
   (Open XML SDK validator, xUnit).
3. **`.cs` and `.resx` files are UTF-8 with BOM and CRLF.** Other text files are UTF-8 and CRLF.
4. **Warnings are errors** (`Directory.Build.props`). Every public member has an XML doc comment; a `<summary>` is
   always written on several lines.
5. **Culture-independent output**: numbers, dates and names written to files use `CultureInfo.InvariantCulture`. User
   messages go through `Messages.Get` / `Messages.Format` and the resources.
6. **Write only.** No reading or round-trip editing of existing documents; callers rebuild documents from their data.
7. **Elements follow the order of ECMA-376.** Excel "repairs" a file in the wrong order; OnlyOffice and LibreOffice
   misread it silently.
8. **Every feature has a test** that saves a file and validates it with the Open XML SDK; add the file to the gallery
   (`XlsxTestKit.Keep`) when a visual check makes sense. Output must be byte-identical for identical input when
   `XlsxWorkbook.Created` is set.
9. **Both targets, one code base.** Every change builds and passes on net48 and net10.0 in Release; .NET-only speed-ups
   go behind `#if NET` with the same output. Refactors keep the gallery files byte-identical, checked per target.
10. **Speed is measured, not assumed**: a change on the save path is benchmarked before and after on both targets.

## Build

```bash
dotnet build Divit.sln
dotnet test Divit.sln
```
