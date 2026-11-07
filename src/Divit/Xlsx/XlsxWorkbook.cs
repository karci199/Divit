#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

/// <summary>
/// An Excel workbook (.xlsx): collects sheets and writes them to disk with <see cref="Save"/>.
/// </summary>
/// <remarks>
/// No external library. An xlsx file is a zip of XML parts (ECMA-376); .NET's own
/// <c>System.IO.Compression</c> and <c>XmlWriter</c> are enough. Parts and the elements inside them are written in the
/// order the standard requires: Excel offers to "repair" a file in the wrong order, OnlyOffice and LibreOffice misread it silently.
/// <para>Text goes into the shared string table (sharedStrings), so each distinct text is stored once. Styles are made
/// unique the same way: a thousand cells that look alike point at one style record.</para>
/// <para>The file is first written next to the target as <c>.tmp</c> and moved into place when complete, so an
/// interrupted save never leaves a half-written file.</para>
/// </remarks>
public sealed partial class XlsxWorkbook
{
    #region Fields

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private readonly List<XlsxSheet> _sheets = new List<XlsxSheet>();

    /// <summary>
    /// Defined names in the order they were added.
    /// </summary>
    private readonly List<(string Name, XlsxSheet Sheet, int Row1, int Col1, int Row2, int Col2, XlsxSheet? Scope)> _names =
        new List<(string, XlsxSheet, int, int, int, int, XlsxSheet?)>();

    /// <summary>
    /// Number of drawing layers and charts written by the last save; the content type list is built from them.
    /// </summary>
    private int _drawings;

    private int _charts;

    /// <summary>
    /// Number of sheets with comments in the last save, and the tab index of the sheet that opens first.
    /// </summary>
    private int _notes;

    private int _active;

    /// <summary>
    /// Whether the last save wrote a formula; set while the sheets are written.
    /// </summary>
    private bool _formulas;

    #endregion

    #region Properties

    /// <summary>
    /// Default font of the workbook; cells whose style names no font use it. The column width unit is also
    /// computed from it.
    /// </summary>
    public string DefaultFontName { get; set; } = "Segoe UI";

    /// <summary>
    /// Default font size in points.
    /// </summary>
    public double DefaultFontSize { get; set; } = 10;

    /// <summary>
    /// Height (pt) of a row without an explicit height: font size × 1.5. Excel uses the same value (customHeight).
    /// </summary>
    public double DefaultRowHeight => DefaultFontSize * 1.5;

    /// <summary>
    /// Application name written into the file properties.
    /// </summary>
    public string Application { get; set; } = "Divit";

    /// <summary>
    /// Opening zoom of sheets without their own zoom, percent (10–400).
    /// </summary>
    public int DefaultZoom { get; set; } = 100;

    /// <summary>
    /// The sheets in tab order.
    /// </summary>
    public IReadOnlyList<XlsxSheet> Sheets => _sheets;

    /// <summary>
    /// File property: title. Shown in Windows Explorer and on Excel's File → Info screen, and searchable.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// File property: subject.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// File property: keywords (for example a serial number, so a search in Explorer finds the file).
    /// </summary>
    public string? Keywords { get; set; }

    /// <summary>
    /// File property: description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// File property: author; <see cref="Application"/> when empty.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// File property: creation and modification time (UTC); the time of saving when empty. Set it to get byte-identical
    /// files from identical input.
    /// </summary>
    public DateTime? Created { get; set; }

    /// <summary>
    /// The sheet that is open when the file opens; the first visible sheet when empty. It must be visible.
    /// </summary>
    public XlsxSheet? ActiveSheet { get; set; }

    #endregion

    #region Public methods

    /// <summary>
    /// Adds a new sheet at the end.
    /// </summary>
    /// <param name="name">Tab name: 1–31 characters, no <c>: \ / ? * [ ]</c>, unique in the workbook (ignoring case).</param>
    /// <returns>The new sheet.</returns>
    public XlsxSheet AddSheet(string name)
    {
        var problem = SheetNameProblem(name);
        if (problem != null)
        {
            throw new ArgumentException(problem, nameof(name));
        }

        var sheet = new XlsxSheet(name);
        _sheets.Add(sheet);
        return sheet;
    }

    /// <summary>
    /// Moves a sheet to another place in the tab order.
    /// </summary>
    /// <param name="sheet">A sheet of this workbook.</param>
    /// <param name="index">New position, 0 for the first tab.</param>
    /// <exception cref="ArgumentException">The sheet belongs to another workbook.</exception>
    public void MoveSheet(XlsxSheet sheet, int index)
    {
        if (!_sheets.Remove(sheet))
        {
            throw new ArgumentException(Messages.Format(Strings.SheetNotInWorkbook, sheet.Name), nameof(sheet));
        }

        _sheets.Insert(Math.Max(0, Math.Min(index, _sheets.Count)), sheet);
    }

    /// <summary>
    /// Gives a range a name, usable in formulas (<c>=SUM(Sales)</c>) and in Excel's Name Box to jump to it.
    /// </summary>
    /// <param name="name">The name: starts with a letter or underscore, then letters, digits, underscores and dots;
    /// not a cell address such as <c>A1</c> or <c>R1C1</c>; unique in its scope (ignoring case).</param>
    /// <param name="sheet">Sheet of the range.</param>
    /// <param name="row1">First row (1-based).</param>
    /// <param name="column1">First column (1-based).</param>
    /// <param name="row2">Last row.</param>
    /// <param name="column2">Last column.</param>
    /// <param name="scope">The sheet the name belongs to, so the same name can exist on other sheets; the whole
    /// workbook when empty.</param>
    /// <exception cref="ArgumentException">The name is not usable or the sheet belongs to another workbook.</exception>
    public void DefineName(string name, XlsxSheet sheet, int row1, int column1, int row2, int column2, XlsxSheet? scope = null)
    {
        foreach (var owned in new[] { sheet, scope })
        {
            if (owned != null && !_sheets.Contains(owned))
            {
                throw new ArgumentException(Messages.Format(Strings.SheetNotInWorkbook, owned.Name), nameof(sheet));
            }
        }

        if (!IsDefinedName(name))
        {
            throw new ArgumentException(Messages.Format(Strings.DefinedNameInvalid, name), nameof(name));
        }

        foreach (var defined in _names)
        {
            if (defined.Scope == scope && string.Equals(defined.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(Messages.Format(Strings.DefinedNameDuplicate, name), nameof(name));
            }
        }

        if (row1 < 1 || column1 < 1 || row2 > XlsxSheet.MaxRow || column2 > XlsxSheet.MaxColumn || row2 < row1 || column2 < column1)
        {
            throw new ArgumentOutOfRangeException(nameof(row1), Messages.Format(Strings.CellOutOfRange, row1, column1));
        }

        _names.Add((name, sheet, row1, column1, row2, column2, scope));
    }

    /// <summary>
    /// Whether a name can be a sheet name; the reason when it cannot, <c>null</c> when it can.
    /// </summary>
    public string? SheetNameProblem(string name)
    {
        if (name.Length == 0 || name.Length > 31)
        {
            return Messages.Format(Strings.SheetNameLength, name);
        }

        if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0 || name[0] == '\'' || name[name.Length - 1] == '\'')
        {
            return Messages.Format(Strings.SheetNameCharacters, name);
        }

        foreach (var sheet in _sheets)
        {
            if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return Messages.Format(Strings.SheetNameDuplicate, name);
            }
        }

        return null;
    }

    /// <summary>
    /// Writes the workbook; an existing target is overwritten.
    /// </summary>
    /// <param name="path">Target file.</param>
    public void Save(string path)
    {
        if (_sheets.Count == 0)
        {
            throw new InvalidOperationException(Messages.Get(Strings.WorkbookEmpty));
        }

        var active = ActiveSheet ?? _sheets.FirstOrDefault(s => s.Visibility == XlsxSheetVisibility.Visible);
        if (active == null)
        {
            throw new InvalidOperationException(Messages.Get(Strings.NoVisibleSheet));
        }

        if (!_sheets.Contains(active))
        {
            throw new InvalidOperationException(Messages.Format(Strings.SheetNotInWorkbook, active.Name));
        }

        if (active.Visibility != XlsxSheetVisibility.Visible)
        {
            throw new InvalidOperationException(Messages.Format(Strings.ActiveSheetHidden, active.Name));
        }

        _active = _sheets.IndexOf(active);
        _formulas = false;
        var plans = PlanSheets();
        var styles = new StyleTable(this);
        var strings = new StringTable();
        OpcPackage.Save(path, Created, package =>
        {
            // Sheets first: the style and string tables fill while they are written.
            for (var i = 0; i < _sheets.Count; i++)
            {
                package.XmlPart("xl/worksheets/sheet" + (i + 1) + ".xml", w => WriteSheet(w, _sheets[i], i == _active, plans[i], styles, strings));
            }

            // The sheets' relationships, then their notes: a comment list and the legacy VML shapes Excel draws them with.
            for (var i = 0; i < _sheets.Count; i++)
            {
                var plan = plans[i];
                package.Relationships("xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels", plan.Relationships);
                if (plan.Comments > 0)
                {
                    var sheet = _sheets[i];
                    package.XmlPart("xl/comments" + plan.Comments + ".xml", w => WriteComments(w, sheet));
                    package.XmlPart("xl/drawings/vmlDrawing" + plan.Comments + ".vml", w => WriteNoteShapes(w, sheet, plan.Comments), false);
                }
            }

            // Every sheet with charts or images: sheet → drawing layer → charts and images. Numbers run across the workbook.
            _charts = 0;
            var media = 0;
            for (var i = 0; i < _sheets.Count; i++)
            {
                var charts = _sheets[i].Charts;
                var images = _sheets[i].Images;
                if (plans[i].Drawing == 0)
                {
                    continue;
                }

                var drawing = plans[i].Drawing;
                var links = new RelationshipList();
                foreach (var image in images)
                {
                    package.BinaryPart("xl/media/image" + (++media) + ".png", image.Png, false);
                }

                for (var k = 0; k < charts.Count; k++)
                {
                    links.Add(RelationshipTypes.Chart, "../charts/chart" + (_charts + 1 + k) + ".xml");
                }

                for (var k = 0; k < images.Count; k++)
                {
                    links.Add(RelationshipTypes.Image, "../media/image" + (media - images.Count + 1 + k) + ".png");
                }

                package.XmlPart("xl/drawings/drawing" + drawing + ".xml", w => XlsxChartWriter.WriteDrawing(w, charts, images));
                package.Relationships("xl/drawings/_rels/drawing" + drawing + ".xml.rels", links);
                foreach (var chart in charts)
                {
                    var number = ++_charts;
                    package.XmlPart("xl/charts/chart" + number + ".xml", w => XlsxChartWriter.WriteChart(w, chart, DefaultFontName));
                }
            }

            package.XmlPart("xl/styles.xml", styles.Write);
            package.XmlPart("xl/sharedStrings.xml", strings.Write);
            package.XmlPart("xl/workbook.xml", WriteWorkbook);
            package.Relationships("xl/_rels/workbook.xml.rels", WorkbookRelationships());
            package.DocumentProperties(new DocumentProperties(Application)
            {
                Title = Title,
                Subject = Subject,
                Keywords = Keywords,
                Description = Description,
                Author = Author,
                Created = Created,
            });

            var root = new RelationshipList();
            root.Add(RelationshipTypes.OfficeDocument, "xl/workbook.xml");
            root.Add(RelationshipTypes.CoreProperties, "docProps/core.xml");
            root.Add(RelationshipTypes.ExtendedProperties, "docProps/app.xml");
            package.Relationships("_rels/.rels", root);
            AddContentTypes(package);
        });
    }

    /// <summary>
    /// Letters of a column number: 1 → A, 27 → AA.
    /// </summary>
    public static string ColumnName(int column)
    {
        var name = string.Empty;
        while (column > 0)
        {
            var rest = (column - 1) % 26;
            name = (char)('A' + rest) + name;
            column = (column - 1) / 26;
        }

        return name;
    }

    /// <summary>
    /// Cell address: (3, 2) → B3.
    /// </summary>
    public static string CellName(int row, int column) => ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

    #endregion

    #region Helpers

    private static string Int(int value) => OfficeText.Int(value);

    /// <summary>
    /// Excel's rules for a defined name: letter, underscore or backslash first, then letters, digits, underscores, dots
    /// and backslashes; at most 255 characters; not a cell address in either notation; not a reserved name.
    /// </summary>
    private static bool IsDefinedName(string name)
    {
        if (name.Length == 0 || name.Length > 255 || name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!(char.IsLetter(name[0]) || name[0] == '_' || name[0] == '\\'))
        {
            return false;
        }

        foreach (var ch in name)
        {
            if (!(char.IsLetterOrDigit(ch) || ch == '_' || ch == '.' || ch == '\\'))
            {
                return false;
            }
        }

        return !LooksLikeA1(name) && !LooksLikeR1C1(name);
    }

    /// <summary>
    /// Letters of an existing column (up to XFD) followed by digits: <c>B3</c>, <c>XFD1</c>.
    /// </summary>
    private static bool LooksLikeA1(string name)
    {
        var letters = 0;
        var column = 0;
        while (letters < name.Length && letters < 4 && IsAsciiLetter(name[letters]))
        {
            column = (column * 26) + (char.ToUpperInvariant(name[letters]) - 'A' + 1);
            letters++;
        }

        if (letters == 0 || letters == name.Length || column > XlsxSheet.MaxColumn)
        {
            return false;
        }

        for (var i = letters; i < name.Length; i++)
        {
            if (name[i] < '0' || name[i] > '9')
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// <c>R</c>, <c>C</c>, <c>R1</c>, <c>C2</c>, <c>R1C2</c> and the like, in any case.
    /// </summary>
    private static bool LooksLikeR1C1(string name)
    {
        var i = 0;
        var parts = 0;
        foreach (var letter in new[] { 'R', 'C' })
        {
            if (i < name.Length && char.ToUpperInvariant(name[i]) == letter)
            {
                i++;
                parts++;
                while (i < name.Length && name[i] >= '0' && name[i] <= '9')
                {
                    i++;
                }
            }
        }

        return parts > 0 && i == name.Length;
    }

    private static bool IsAsciiLetter(char ch) => (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z');

    private static void Cfvo(XmlWriter w, string type, string? value)
    {
        w.WriteStartElement("cfvo", MainNs);
        w.WriteAttributeString("type", type);
        if (value != null) { w.WriteAttributeString("val", value); }
        w.WriteEndElement();
    }

    /// <summary>
    /// Excel's legacy 16-bit password hash for sheet protection (ECMA-376, part 18.2.28); four hex digits.
    /// </summary>
    internal static string PasswordHash(string password)
    {
        var hash = 0;
        for (var i = password.Length - 1; i >= 0; i--)
        {
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= password[i];
        }

        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= password.Length;
        hash ^= 0xCE4B;
        return hash.ToString("X4", CultureInfo.InvariantCulture);
    }

    private static string Num(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    /// <summary>
    /// A margin in inches; four decimals keep margins given in centimetres exact to a hundredth of a millimetre.
    /// </summary>
    private static string Inches(double value) => Math.Max(0, value).ToString("0.####", CultureInfo.InvariantCulture);

    /// <summary>
    /// <c>#RRGGBB</c> → Excel's <c>FFRRGGBB</c>.
    /// </summary>
    internal static string Argb(string color) => "FF" + OfficeText.Rgb(color);

    #endregion
}
