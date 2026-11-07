#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

/// <content>
/// The workbook-level parts: part numbering, workbook, its relationships and the content types of the package.
/// </content>
public sealed partial class XlsxWorkbook
{
    #region Package

    /// <summary>
    /// Numbers the parts of every sheet and gives their relationships ids, before anything is written: the sheet XML
    /// names the ids, the relationship parts and content types list the parts.
    /// </summary>
    private SheetPlan[] PlanSheets()
    {
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
                plan.DrawingId = plan.Relationships.Add(RelationshipTypes.Drawing, "../drawings/drawing" + plan.Drawing + ".xml");
            }

            if (sheet.Comments.Count > 0)
            {
                plan.Comments = ++_notes;
                plan.Relationships.Add(RelationshipTypes.Comments, "../comments" + plan.Comments + ".xml");
                plan.NoteShapesId = plan.Relationships.Add(RelationshipTypes.VmlDrawing, "../drawings/vmlDrawing" + plan.Comments + ".vml");
            }

            foreach (var link in sheet.UrlLinks)
            {
                plan.LinkIds.Add(plan.Relationships.Add(RelationshipTypes.Hyperlink, link.Url, true));
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

    /// <summary>
    /// The workbook's relationships: the sheets as rId1 … rIdN, then styles and shared strings.
    /// </summary>
    private RelationshipList WorkbookRelationships()
    {
        var list = new RelationshipList();
        for (var i = 0; i < _sheets.Count; i++)
        {
            list.Add(RelationshipTypes.Worksheet, "worksheets/sheet" + (i + 1) + ".xml");
        }

        list.Add(RelationshipTypes.Styles, "styles.xml");
        list.Add(RelationshipTypes.SharedStrings, "sharedStrings.xml");
        return list;
    }

    /// <summary>
    /// Content types: extensions first, then the parts in the order of the workbook.
    /// </summary>
    private void AddContentTypes(OpcPackage package)
    {
        package.Default("rels", ContentTypes.Relationships);
        package.Default("xml", ContentTypes.Xml);
        package.Default("png", ContentTypes.Png);
        if (_notes > 0)
        {
            package.Default("vml", ContentTypes.Vml);
        }

        package.Override("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        for (var i = 0; i < _sheets.Count; i++)
        {
            package.Override("/xl/worksheets/sheet" + (i + 1) + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
        }

        for (var i = 1; i <= _drawings; i++)
        {
            package.Override("/xl/drawings/drawing" + i + ".xml", ContentTypes.Drawing);
        }

        for (var i = 1; i <= _charts; i++)
        {
            package.Override("/xl/charts/chart" + i + ".xml", ContentTypes.Chart);
        }

        for (var i = 1; i <= _notes; i++)
        {
            package.Override("/xl/comments" + i + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
        }

        package.Override("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
        package.Override("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
        package.Override("/docProps/core.xml", ContentTypes.CoreProperties);
        package.Override("/docProps/app.xml", ContentTypes.ExtendedProperties);
    }

    #endregion

    /// <summary>
    /// Parts and relationship ids of one sheet; numbers 0 mean the sheet has no such part.
    /// </summary>
    private sealed class SheetPlan
    {
        public readonly RelationshipList Relationships = new RelationshipList();

        /// <summary>
        /// Relationship ids of the URL links, in the order of <see cref="XlsxSheet.UrlLinks"/>.
        /// </summary>
        public readonly List<string> LinkIds = new List<string>();

        public int Drawing;

        public int Comments;

        public string? DrawingId;

        public string? NoteShapesId;
    }
}
