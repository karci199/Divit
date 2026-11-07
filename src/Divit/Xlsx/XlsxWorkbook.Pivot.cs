#nullable enable
namespace Divit.Xlsx;

using Divit.Packaging;
using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Xml;

/// <content>
/// Pivot tables: the cache (definition and records) and the table definition.
/// </content>
public sealed partial class XlsxWorkbook
{
    private const string PivotTableRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    private const string PivotCacheRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    private const string PivotRecordsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
    private const string PivotTableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
    private const string PivotCacheContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
    private const string PivotRecordsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

    /// <summary>
    /// The pivot tables of the last save, in sheet order; table i uses cache i.
    /// </summary>
    private readonly List<XlsxPivotTable> _pivots = new List<XlsxPivotTable>();

    /// <summary>
    /// Numbers the pivot tables, checks their names and sources, and writes their summaries into the cells.
    /// </summary>
    /// <exception cref="InvalidOperationException">Two pivot tables on a sheet share a name, a source is streamed or
    /// belongs to another workbook, or a field is wrong.</exception>
    private void PreparePivots()
    {
        _pivots.Clear();
        foreach (var sheet in _sheets)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pivot in sheet.PivotTables)
            {
                if (!names.Add(pivot.Name))
                {
                    throw new InvalidOperationException(Messages.Format(Strings.PivotNameDuplicate, pivot.Name));
                }

                if (!_sheets.Contains(pivot.Source))
                {
                    throw new InvalidOperationException(Messages.Format(Strings.SheetNotInWorkbook, pivot.Source.Name));
                }

                if (pivot.Source.IsStreamed)
                {
                    throw new InvalidOperationException(Messages.Format(Strings.SheetIsStreamed, pivot.Source.Name));
                }

                _pivots.Add(pivot);
                pivot.Id = _pivots.Count;
                pivot.Prepare();
            }
        }
    }

    /// <summary>
    /// The pivot caches of the workbook, named in workbook.xml after the calculation settings.
    /// </summary>
    private void WritePivotCaches(XmlWriter w)
    {
        if (_pivots.Count == 0)
        {
            return;
        }

        w.WriteStartElement("pivotCaches", MainNs);
        foreach (var pivot in _pivots)
        {
            w.WriteStartElement("pivotCache", MainNs);
            w.WriteAttributeString("cacheId", Int(pivot.Id));
            w.WriteAttributeString("id", RelNs, PivotCacheId(pivot));
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// Relationship id of a cache in the workbook relationships: after the sheets, styles, shared strings and theme.
    /// </summary>
    private string PivotCacheId(XlsxPivotTable pivot) => "rId" + Int(_sheets.Count + 3 + pivot.Id);

    /// <summary>
    /// The cache definition: where the data comes from and, per field, its distinct items (fields on an axis or
    /// holding text) or the range of its numbers. Excel refreshes the cache when the file opens.
    /// </summary>
    private static void WritePivotCache(XmlWriter w, XlsxPivotTable pivot)
    {
        w.WriteStartElement("pivotCacheDefinition", MainNs);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        w.WriteAttributeString("id", RelNs, "rId1");
        w.WriteAttributeString("refreshOnLoad", "1");
        w.WriteAttributeString("createdVersion", "6");
        w.WriteAttributeString("refreshedVersion", "6");
        w.WriteAttributeString("minRefreshableVersion", "3");
        w.WriteAttributeString("recordCount", Int(pivot.RecordCount));
        w.WriteStartElement("cacheSource", MainNs);
        w.WriteAttributeString("type", "worksheet");
        w.WriteStartElement("worksheetSource", MainNs);
        w.WriteAttributeString("ref", CellName(pivot.SourceRow1, pivot.SourceColumn1) + ":" + CellName(pivot.SourceRow2, pivot.SourceColumn2));
        w.WriteAttributeString("sheet", pivot.Source.Name);
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("cacheFields", MainNs);
        w.WriteAttributeString("count", Int(pivot.Fields.Length));
        foreach (var field in pivot.Fields)
        {
            w.WriteStartElement("cacheField", MainNs);
            w.WriteAttributeString("name", OfficeText.Clean(field.Name));
            w.WriteAttributeString("numFmtId", "0");
            w.WriteStartElement("sharedItems", MainNs);
            SharedItemFlags(w, field);
            if (Listed(field))
            {
                w.WriteAttributeString("count", Int(field.Items.Count));
                foreach (var item in field.Items)
                {
                    Item(w, item);
                }
            }

            w.WriteEndElement();
            w.WriteEndElement();
        }

        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// Whether a field lists its items in the cache (and its records point at them) instead of holding numbers.
    /// </summary>
    private static bool Listed(XlsxPivotTable.Field field) => field.OnAxis || !field.Numeric;

    /// <summary>
    /// What kinds of values a field holds; Excel checks these against the items and records.
    /// </summary>
    private static void SharedItemFlags(XmlWriter w, XlsxPivotTable.Field field)
    {
        if (field.HasNumber)
        {
            var min = double.MaxValue;
            var max = double.MinValue;
            var integers = true;
            foreach (var number in field.Numbers)
            {
                if (number is double v)
                {
                    min = Math.Min(min, v);
                    max = Math.Max(max, v);
                    integers &= v == Math.Floor(v) && Math.Abs(v) < 1e15;
                }
            }

            if (field.HasText)
            {
                w.WriteAttributeString("containsMixedTypes", "1");
            }
            else
            {
                w.WriteAttributeString("containsSemiMixedTypes", "0");
                w.WriteAttributeString("containsString", "0");
            }

            w.WriteAttributeString("containsNumber", "1");
            if (integers) { w.WriteAttributeString("containsInteger", "1"); }
            w.WriteAttributeString("minValue", XlsxRule.Number(min));
            w.WriteAttributeString("maxValue", XlsxRule.Number(max));
        }

        if (field.HasBlank)
        {
            w.WriteAttributeString("containsBlank", "1");
        }
    }

    private static void Item(XmlWriter w, object? item)
    {
        switch (item)
        {
            case null:
                w.WriteElementString("m", MainNs, null);
                break;
            case double number:
                w.WriteStartElement("n", MainNs);
                w.WriteAttributeString("v", XlsxRule.Number(number));
                w.WriteEndElement();
                break;
            default:
                w.WriteStartElement("s", MainNs);
                w.WriteAttributeString("v", OfficeText.Clean((string)item));
                w.WriteEndElement();
                break;
        }
    }

    /// <summary>
    /// The records: per source row, the item index of each listed field and the number of each other field.
    /// </summary>
    private static void WritePivotRecords(XmlWriter w, XlsxPivotTable pivot)
    {
        var buffer = new char[64];
        w.WriteStartElement("pivotCacheRecords", MainNs);
        w.WriteAttributeString("xmlns", "r", null, RelNs);
        w.WriteAttributeString("count", Int(pivot.RecordCount));
        for (var r = 0; r < pivot.RecordCount; r++)
        {
            w.WriteStartElement("r", MainNs);
            foreach (var field in pivot.Fields)
            {
                if (Listed(field))
                {
                    w.WriteStartElement("x", MainNs);
                    WriteNumberAttribute(w, "v", field.Records[r], buffer);
                    w.WriteEndElement();
                }
                else if (field.Numbers[r] is double number)
                {
                    w.WriteStartElement("n", MainNs);
                    w.WriteStartAttribute("v");
                    w.WriteChars(buffer, 0, OfficeText.FormatDouble(buffer, number));
                    w.WriteEndAttribute();
                    w.WriteEndElement();
                }
                else
                {
                    w.WriteElementString("m", MainNs, null);
                }
            }

            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    /// <summary>
    /// The table definition in Excel's compact layout: where it sits, the role of each field, the row and column items
    /// in display order, the filter, and the value field with its function.
    /// </summary>
    private static void WritePivotTable(XmlWriter w, XlsxPivotTable pivot, StyleTable styles)
    {
        var columns = pivot.ColumnIndex >= 0;
        w.WriteStartElement("pivotTableDefinition", MainNs);
        w.WriteAttributeString("name", OfficeText.Clean(pivot.Name));
        w.WriteAttributeString("cacheId", Int(pivot.Id));
        w.WriteAttributeString("applyNumberFormats", "0");
        w.WriteAttributeString("applyBorderFormats", "0");
        w.WriteAttributeString("applyFontFormats", "0");
        w.WriteAttributeString("applyPatternFormats", "0");
        w.WriteAttributeString("applyAlignmentFormats", "0");
        w.WriteAttributeString("applyWidthHeightFormats", "1");
        w.WriteAttributeString("dataCaption", "Values");
        w.WriteAttributeString("updatedVersion", "6");
        w.WriteAttributeString("minRefreshableVersion", "3");
        w.WriteAttributeString("useAutoFormatting", "1");
        w.WriteAttributeString("itemPrintTitles", "1");
        w.WriteAttributeString("createdVersion", "6");
        w.WriteAttributeString("indent", "0");
        w.WriteAttributeString("outline", "1");
        w.WriteAttributeString("outlineData", "1");
        w.WriteAttributeString("multipleFieldFilters", "0");

        w.WriteStartElement("location", MainNs);
        w.WriteAttributeString("ref", CellName(pivot.BodyRow, pivot.Column) + ":" + CellName(pivot.LastRow, pivot.LastColumn));
        w.WriteAttributeString("firstHeaderRow", "1");
        w.WriteAttributeString("firstDataRow", columns ? "2" : "1");
        w.WriteAttributeString("firstDataCol", "1");
        if (pivot.FilterIndex >= 0)
        {
            w.WriteAttributeString("rowPageCount", "1");
            w.WriteAttributeString("colPageCount", "1");
        }

        w.WriteEndElement();

        w.WriteStartElement("pivotFields", MainNs);
        w.WriteAttributeString("count", Int(pivot.Fields.Length));
        for (var i = 0; i < pivot.Fields.Length; i++)
        {
            w.WriteStartElement("pivotField", MainNs);
            var axis = i == pivot.RowIndex ? "axisRow" : i == pivot.ColumnIndex ? "axisCol" : i == pivot.FilterIndex ? "axisPage" : null;
            if (axis != null) { w.WriteAttributeString("axis", axis); }
            if (i == pivot.ValueIndex) { w.WriteAttributeString("dataField", "1"); }
            w.WriteAttributeString("showAll", "0");
            if (axis != null)
            {
                var order = i == pivot.RowIndex ? pivot.RowOrder : i == pivot.ColumnIndex ? pivot.ColumnOrder : pivot.Fields[i].SortedItems();
                w.WriteStartElement("items", MainNs);
                w.WriteAttributeString("count", Int(order.Length + 1));
                foreach (var item in order)
                {
                    w.WriteStartElement("item", MainNs);
                    w.WriteAttributeString("x", Int(item));
                    w.WriteEndElement();
                }

                w.WriteStartElement("item", MainNs);
                w.WriteAttributeString("t", "default");
                w.WriteEndElement();
                w.WriteEndElement();
            }

            w.WriteEndElement();
        }

        w.WriteEndElement();

        AxisItems(w, "rowFields", "rowItems", pivot.RowIndex, pivot.RowOrder.Length);
        if (columns)
        {
            AxisItems(w, "colFields", "colItems", pivot.ColumnIndex, pivot.ColumnOrder.Length);
        }
        else
        {
            // One column of values: a single, empty column item.
            w.WriteStartElement("colItems", MainNs);
            w.WriteAttributeString("count", "1");
            w.WriteElementString("i", MainNs, null);
            w.WriteEndElement();
        }

        if (pivot.FilterIndex >= 0)
        {
            w.WriteStartElement("pageFields", MainNs);
            w.WriteAttributeString("count", "1");
            w.WriteStartElement("pageField", MainNs);
            w.WriteAttributeString("fld", Int(pivot.FilterIndex));
            w.WriteAttributeString("hier", "-1");
            w.WriteEndElement();
            w.WriteEndElement();
        }

        w.WriteStartElement("dataFields", MainNs);
        w.WriteAttributeString("count", "1");
        w.WriteStartElement("dataField", MainNs);
        w.WriteAttributeString("name", OfficeText.Clean(pivot.ValueCaption));
        w.WriteAttributeString("fld", Int(pivot.ValueIndex));
        switch (pivot.Function)
        {
            case XlsxPivotFunction.Count: w.WriteAttributeString("subtotal", "count"); break;
            case XlsxPivotFunction.Average: w.WriteAttributeString("subtotal", "average"); break;
            case XlsxPivotFunction.Max: w.WriteAttributeString("subtotal", "max"); break;
            case XlsxPivotFunction.Min: w.WriteAttributeString("subtotal", "min"); break;
        }

        w.WriteAttributeString("baseField", "0");
        w.WriteAttributeString("baseItem", "0");
        if (pivot.NumberFormat != null)
        {
            w.WriteAttributeString("numFmtId", Int(styles.NumberFormatId(pivot.NumberFormat)));
        }

        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement("pivotTableStyleInfo", MainNs);
        w.WriteAttributeString("name", pivot.Style);
        w.WriteAttributeString("showRowHeaders", "1");
        w.WriteAttributeString("showColHeaders", "1");
        w.WriteAttributeString("showRowStripes", "0");
        w.WriteAttributeString("showColStripes", "0");
        w.WriteAttributeString("showLastColumn", "1");
        w.WriteEndElement();
        w.WriteEndElement();
    }

    /// <summary>
    /// The field of an axis and its items: one per display position, then the grand total.
    /// </summary>
    private static void AxisItems(XmlWriter w, string fieldsName, string itemsName, int field, int count)
    {
        w.WriteStartElement(fieldsName, MainNs);
        w.WriteAttributeString("count", "1");
        w.WriteStartElement("field", MainNs);
        w.WriteAttributeString("x", Int(field));
        w.WriteEndElement();
        w.WriteEndElement();

        w.WriteStartElement(itemsName, MainNs);
        w.WriteAttributeString("count", Int(count + 1));
        for (var i = 0; i < count; i++)
        {
            w.WriteStartElement("i", MainNs);
            w.WriteStartElement("x", MainNs);
            if (i > 0) { w.WriteAttributeString("v", Int(i)); }
            w.WriteEndElement();
            w.WriteEndElement();
        }

        w.WriteStartElement("i", MainNs);
        w.WriteAttributeString("t", "grand");
        w.WriteElementString("x", MainNs, null);
        w.WriteEndElement();
        w.WriteEndElement();
    }
}
