#nullable enable
namespace Divit.Xlsx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

/// <content>
/// The package: zip entries, part numbering and the workbook-level parts (workbook, relationships, content types,
/// document properties).
/// </content>
public sealed partial class XlsxWorkbook
{
    #region Parts

    private delegate void PartWriter(XmlWriter w);

    /// <summary>
    /// A zip entry stamped with <see cref="Created"/>: without it the entry carries the clock time of the save and two
    /// saves of the same workbook differ in bytes. Zip times run from 1980 to 2107; outside that the clock is used.
    /// </summary>
    private ZipArchiveEntry Entry(ZipArchive zip, string name, CompressionLevel level)
    {
        var entry = zip.CreateEntry(name, level);
        if (Created is DateTime created)
        {
            var utc = created.ToUniversalTime();
            if (utc.Year >= 1980 && utc.Year <= 2107)
            {
                entry.LastWriteTime = new DateTimeOffset(DateTime.SpecifyKind(utc, DateTimeKind.Unspecified), TimeSpan.Zero);
            }
        }

        return entry;
    }

    /// <summary>
    /// Writes one XML part; <paramref name="declaration"/> is off only for the legacy VML part, which Excel writes
    /// without one.
    /// </summary>
    private void Part(ZipArchive zip, string name, PartWriter write, bool declaration = true)
    {
        var entry = Entry(zip, name, CompressionLevel.Fastest);
        var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, CloseOutput = true, OmitXmlDeclaration = !declaration };
        using (var w = XmlWriter.Create(entry.Open(), settings))
        {
            if (declaration)
            {
                w.WriteStartDocument(true);
            }

            write(w);
            w.WriteEndDocument();
        }
    }

    /// <summary>
    /// Numbers the parts of every sheet and gives their relationships ids, before anything is written: the sheet XML
    /// names the ids, the relationship parts and content types list the parts.
    /// </summary>
    private SheetPlan[] PlanSheets()
    {
        const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        var plans = new SheetPlan[_sheets.Count];
        _drawings = 0;
        _notes = 0;
        for (var i = 0; i < _sheets.Count; i++)
        {
            var sheet = _sheets[i];
            var plan = new SheetPlan();
            if (sheet.Charts.Count > 0 || sheet.Images.Count > 0)
            {
                plan.Drawing = ++_drawings;
                plan.DrawingId = plan.Add(Rel + "drawing", "../drawings/drawing" + plan.Drawing + ".xml", false);
            }

            if (sheet.Comments.Count > 0)
            {
                plan.Comments = ++_notes;
                plan.Add(Rel + "comments", "../comments" + plan.Comments + ".xml", false);
                plan.NoteShapesId = plan.Add(Rel + "vmlDrawing", "../drawings/vmlDrawing" + plan.Comments + ".vml", false);
            }

            foreach (var link in sheet.UrlLinks)
            {
                plan.LinkIds.Add(plan.Add(Rel + "hyperlink", link.Url, true));
            }

            plans[i] = plan;
        }

        return plans;
    }

    private void WriteWorkbook(XmlWriter w)
    {
        w.WriteStartElement("workbook", MainNs);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        w.WriteStartElement("bookViews", MainNs);
        // firstSheet is the leftmost tab shown in the tab bar; a hidden one there leaves the bar looking empty.
        w.WriteStartElement("workbookView", MainNs);
        var first = _sheets.FindIndex(s => s.Visibility == XlsxSheetVisibility.Visible);
        if (first > 0) { w.WriteAttributeString("firstSheet", Int(first)); }
        w.WriteAttributeString("activeTab", Int(_active));
        w.WriteEndElement();
        w.WriteEndElement();
        w.WriteStartElement("sheets", MainNs);
        for (var i = 0; i < _sheets.Count; i++)
        {
            w.WriteStartElement("sheet", MainNs);
            w.WriteAttributeString("name", _sheets[i].Name);
            w.WriteAttributeString("sheetId", Int(i + 1));
            if (_sheets[i].Visibility == XlsxSheetVisibility.Hidden) { w.WriteAttributeString("state", "hidden"); }
            if (_sheets[i].Visibility == XlsxSheetVisibility.VeryHidden) { w.WriteAttributeString("state", "veryHidden"); }
            w.WriteAttributeString("id", RelNs, "rId" + (i + 1));
            w.WriteEndElement();
        }

        w.WriteEndElement();

        // The caller's names, then print area and print titles: reserved names bound to the sheet (_xlnm.Print_Area,
        // _xlnm.Print_Titles).
        var named = _sheets.Where(s => s.PrintTitleRows != null || s.PrintArea != null).ToList();
        if (named.Count > 0 || _names.Count > 0)
        {
            w.WriteStartElement("definedNames", MainNs);
            foreach (var name in _names)
            {
                w.WriteStartElement("definedName", MainNs);
                w.WriteAttributeString("name", name.Name);
                if (name.Scope != null) { w.WriteAttributeString("localSheetId", Int(_sheets.IndexOf(name.Scope))); }
                w.WriteString("'" + name.Sheet.Name.Replace("'", "''") + "'!$" + ColumnName(name.Col1) + "$" + Int(name.Row1)
                    + (name.Row1 == name.Row2 && name.Col1 == name.Col2 ? string.Empty : ":$" + ColumnName(name.Col2) + "$" + Int(name.Row2)));
                w.WriteEndElement();
            }

            foreach (var sheet in named)
            {
                var quoted = "'" + sheet.Name.Replace("'", "''") + "'!";
                if (sheet.PrintArea != null)
                {
                    var area = sheet.PrintArea.Value;
                    w.WriteStartElement("definedName", MainNs);
                    w.WriteAttributeString("name", "_xlnm.Print_Area");
                    w.WriteAttributeString("localSheetId", Int(_sheets.IndexOf(sheet)));
                    w.WriteString(quoted + "$" + ColumnName(area.FirstColumn) + "$" + Int(area.FirstRow) + ":$" + ColumnName(area.LastColumn) + "$" + Int(area.LastRow));
                    w.WriteEndElement();
                }

                if (sheet.PrintTitleRows != null)
                {
                    var rows = sheet.PrintTitleRows.Value;
                    w.WriteStartElement("definedName", MainNs);
                    w.WriteAttributeString("name", "_xlnm.Print_Titles");
                    w.WriteAttributeString("localSheetId", Int(_sheets.IndexOf(sheet)));
                    w.WriteString(quoted + "$" + Int(rows.First) + ":$" + Int(rows.Last));
                    w.WriteEndElement();
                }
            }

            w.WriteEndElement();
        }

        // Formulas carry no trusted results: Excel computes every one of them when the file opens.
        if (_formulas)
        {
            w.WriteStartElement("calcPr", MainNs);
            w.WriteAttributeString("calcId", "191029");
            w.WriteAttributeString("fullCalcOnLoad", "1");
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private void WriteWorkbookRels(XmlWriter w)
    {
        w.WriteStartElement("Relationships", PackageRelNs);
        for (var i = 0; i < _sheets.Count; i++)
        {
            Relationship(w, "rId" + (i + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "worksheets/sheet" + (i + 1) + ".xml");
        }

        Relationship(w, "rId" + (_sheets.Count + 1), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml");
        Relationship(w, "rId" + (_sheets.Count + 2), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml");
        w.WriteEndElement();
    }

    private static void WriteRootRels(XmlWriter w)
    {
        w.WriteStartElement("Relationships", PackageRelNs);
        Relationship(w, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
        Relationship(w, "rId2", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml");
        Relationship(w, "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml");
        w.WriteEndElement();
    }

    private void WriteContentTypes(XmlWriter w)
    {
        w.WriteStartElement("Types", ContentTypesNs);
        Default(w, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        Default(w, "xml", "application/xml");
        Default(w, "png", "image/png");
        if (_notes > 0)
        {
            Default(w, "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing");
        }

        Override(w, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        for (var i = 0; i < _sheets.Count; i++)
        {
            Override(w, "/xl/worksheets/sheet" + (i + 1) + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        }

        for (var i = 1; i <= _drawings; i++)
        {
            Override(w, "/xl/drawings/drawing" + i + ".xml", "application/vnd.openxmlformats-officedocument.drawing+xml");
        }

        for (var i = 1; i <= _charts; i++)
        {
            Override(w, "/xl/charts/chart" + i + ".xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
        }

        for (var i = 1; i <= _notes; i++)
        {
            Override(w, "/xl/comments" + i + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
        }

        Override(w, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        Override(w, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        Override(w, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
        Override(w, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
        w.WriteEndElement();
    }

    private void WriteCore(XmlWriter w)
    {
        const string Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        const string Dc = "http://purl.org/dc/elements/1.1/";
        const string Dcterms = "http://purl.org/dc/terms/";
        const string Xsi = "http://www.w3.org/2001/XMLSchema-instance";
        var now = (Created ?? DateTime.UtcNow).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        w.WriteStartElement("cp", "coreProperties", Cp);
        w.WriteAttributeString("xmlns", "dc", null, Dc);
        w.WriteAttributeString("xmlns", "dcterms", null, Dcterms);
        w.WriteAttributeString("xmlns", "xsi", null, Xsi);
        if (Title != null) { w.WriteElementString("dc", "title", Dc, Title); }
        if (Subject != null) { w.WriteElementString("dc", "subject", Dc, Subject); }
        w.WriteElementString("dc", "creator", Dc, Author ?? Application);
        if (Keywords != null) { w.WriteElementString("cp", "keywords", Cp, Keywords); }
        if (Description != null) { w.WriteElementString("dc", "description", Dc, Description); }
        foreach (var name in new[] { "created", "modified" })
        {
            w.WriteStartElement("dcterms", name, Dcterms);
            w.WriteAttributeString("xsi", "type", Xsi, "dcterms:W3CDTF");
            w.WriteString(now);
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    private void WriteApp(XmlWriter w)
    {
        w.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        w.WriteElementString("Application", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", Application);
        w.WriteEndElement();
    }

    private static void Relationship(XmlWriter w, string id, string type, string target, bool external = false)
    {
        w.WriteStartElement("Relationship", PackageRelNs);
        w.WriteAttributeString("Id", id);
        w.WriteAttributeString("Type", type);
        w.WriteAttributeString("Target", target);
        if (external) { w.WriteAttributeString("TargetMode", "External"); }
        w.WriteEndElement();
    }

    private static void Default(XmlWriter w, string extension, string type)
    {
        w.WriteStartElement("Default", ContentTypesNs);
        w.WriteAttributeString("Extension", extension);
        w.WriteAttributeString("ContentType", type);
        w.WriteEndElement();
    }

    private static void Override(XmlWriter w, string part, string type)
    {
        w.WriteStartElement("Override", ContentTypesNs);
        w.WriteAttributeString("PartName", part);
        w.WriteAttributeString("ContentType", type);
        w.WriteEndElement();
    }

    #endregion

    /// <summary>
    /// Parts and relationship ids of one sheet; numbers 0 mean the sheet has no such part.
    /// </summary>
    private sealed class SheetPlan
    {
        public readonly List<(string Id, string Type, string Target, bool External)> Relationships =
            new List<(string, string, string, bool)>();

        /// <summary>
        /// Relationship ids of the URL links, in the order of <see cref="XlsxSheet.UrlLinks"/>.
        /// </summary>
        public readonly List<string> LinkIds = new List<string>();

        public int Drawing;

        public int Comments;

        public string? DrawingId;

        public string? NoteShapesId;

        public string Add(string type, string target, bool external)
        {
            var id = "rId" + Int(Relationships.Count + 1);
            Relationships.Add((id, type, target, external));
            return id;
        }
    }
}
