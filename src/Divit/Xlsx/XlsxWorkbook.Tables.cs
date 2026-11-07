#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Xml;

/// <content>
/// Table parts (ListObjects).
/// </content>
public sealed partial class XlsxWorkbook
{
    private const string TableRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    private const string TableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

    /// <summary>
    /// Number of tables in the last save.
    /// </summary>
    private int _tables;

    /// <summary>
    /// Checks the tables against each other and the sheets, numbers them across the workbook and fills their header
    /// and totals cells.
    /// </summary>
    /// <exception cref="InvalidOperationException">Two tables or a table and a defined name share a name, a table
    /// holds merged cells, overlaps the sheet's auto filter or, with its totals row, another table.</exception>
    private void PrepareTables()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var defined in _names)
        {
            names.Add(defined.Name);
        }

        _tables = 0;
        foreach (var sheet in _sheets)
        {
            for (var t = 0; t < sheet.Tables.Count; t++)
            {
                var table = sheet.Tables[t];
                if (!names.Add(table.Name))
                {
                    throw new InvalidOperationException(Messages.Format(Strings.TableNameDuplicate, table.Name));
                }

                foreach (var merge in sheet.Merges)
                {
                    if (Overlaps(table, merge.Row1, merge.Col1, merge.Row2, merge.Col2))
                    {
                        throw new InvalidOperationException(Messages.Format(Strings.TableOverlapsMerge, table.Name));
                    }
                }

                if (sheet.AutoFilterRange is var filter && filter != null
                    && Overlaps(table, filter.Value.Row1, filter.Value.Col1, filter.Value.Row2, filter.Value.Col2))
                {
                    throw new InvalidOperationException(Messages.Format(Strings.TableOverlapsFilter, table.Name));
                }

                // A totals row switched on after adding can run into the table below.
                for (var o = 0; o < t; o++)
                {
                    var other = sheet.Tables[o];
                    if (Overlaps(table, other.Row1, other.Column1, other.LastRow, other.Column2))
                    {
                        throw new InvalidOperationException(Messages.Format(Strings.TablesOverlap, table.Name, other.Name));
                    }
                }

                table.Id = ++_tables;
                table.Prepare();
            }
        }
    }

    private static bool Overlaps(XlsxTable table, int row1, int column1, int row2, int column2) =>
        row1 <= table.LastRow && table.Row1 <= row2 && column1 <= table.Column2 && table.Column1 <= column2;

    /// <summary>
    /// One table part: range, filter, columns with their totals, style.
    /// </summary>
    private static void WriteTable(XmlWriter w, XlsxTable table)
    {
        w.WriteStartElement("table", MainNs);
        w.WriteAttributeString("id", Int(table.Id));
        w.WriteAttributeString("name", table.Name);
        w.WriteAttributeString("displayName", table.Name);
        w.WriteAttributeString("ref", CellName(table.Row1, table.Column1) + ":" + CellName(table.LastRow, table.Column2));
        if (table.TotalsRow)
        {
            w.WriteAttributeString("totalsRowCount", "1");
        }
        else
        {
            w.WriteAttributeString("totalsRowShown", "0");
        }

        // The filter covers header and data, not the totals row.
        if (table.FilterButtons)
        {
            w.WriteStartElement("autoFilter", MainNs);
            w.WriteAttributeString("ref", CellName(table.Row1, table.Column1) + ":" + CellName(table.Row2, table.Column2));
            w.WriteEndElement();
        }

        w.WriteStartElement("tableColumns", MainNs);
        w.WriteAttributeString("count", Int(table.ColumnNames.Count));
        for (var i = 0; i < table.ColumnNames.Count; i++)
        {
            w.WriteStartElement("tableColumn", MainNs);
            w.WriteAttributeString("id", Int(i + 1));
            w.WriteAttributeString("name", Divit.Packaging.OfficeText.Clean(table.ColumnNames[i]));
            if (table.TotalsRow && table.LabelOf(i) is string label)
            {
                w.WriteAttributeString("totalsRowLabel", Divit.Packaging.OfficeText.Clean(label));
            }
            else if (table.TotalsRow && table.TotalOf(i) != XlsxTotal.None)
            {
                w.WriteAttributeString("totalsRowFunction", XlsxTable.FunctionName(table.TotalOf(i)));
            }

            w.WriteEndElement();
        }

        w.WriteEndElement();

        w.WriteStartElement("tableStyleInfo", MainNs);
        if (table.Style != null) { w.WriteAttributeString("name", table.Style); }
        w.WriteAttributeString("showFirstColumn", table.FirstColumn ? "1" : "0");
        w.WriteAttributeString("showLastColumn", table.LastColumn ? "1" : "0");
        w.WriteAttributeString("showRowStripes", table.RowStripes ? "1" : "0");
        w.WriteAttributeString("showColumnStripes", table.ColumnStripes ? "1" : "0");
        w.WriteEndElement();
        w.WriteEndElement();
    }
}
