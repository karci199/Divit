#nullable enable
namespace Divit.Packaging;

using System.Collections.Generic;
using System.Xml;

/// <summary>
/// The relationships of one part (its <c>.rels</c> part): ids <c>rId1</c>, <c>rId2</c> … in the order they are added.
/// </summary>
internal sealed class RelationshipList
{
    private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    private readonly List<(string Id, string Type, string Target, bool External)> _items =
        new List<(string, string, string, bool)>();

    /// <summary>
    /// Number of relationships.
    /// </summary>
    public int Count => _items.Count;

    /// <summary>
    /// Adds a relationship and returns its id.
    /// </summary>
    /// <param name="type">Relationship type, one of <see cref="RelationshipTypes"/>.</param>
    /// <param name="target">Target part relative to the source part's folder, or an address outside the package.</param>
    /// <param name="external">Whether the target lies outside the package (a web address, a file).</param>
    /// <returns>The id, such as <c>rId3</c>.</returns>
    public string Add(string type, string target, bool external = false)
    {
        var id = "rId" + OfficeText.Int(_items.Count + 1);
        _items.Add((id, type, target, external));
        return id;
    }

    /// <summary>
    /// Writes the relationship part.
    /// </summary>
    public void Write(XmlWriter w)
    {
        w.WriteStartElement("Relationships", PackageRelNs);
        foreach (var item in _items)
        {
            w.WriteStartElement("Relationship", PackageRelNs);
            w.WriteAttributeString("Id", item.Id);
            w.WriteAttributeString("Type", item.Type);
            w.WriteAttributeString("Target", item.Target);
            if (item.External) { w.WriteAttributeString("TargetMode", "External"); }
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }
}

/// <summary>
/// Relationship types shared by the Office formats.
/// </summary>
internal static class RelationshipTypes
{
    private const string Office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    /// <summary>
    /// The main part of the document (workbook, document, presentation).
    /// </summary>
    public const string OfficeDocument = Office + "officeDocument";

    /// <summary>
    /// Core properties: title, author, dates.
    /// </summary>
    public const string CoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    /// <summary>
    /// Extended properties: application.
    /// </summary>
    public const string ExtendedProperties = Office + "extended-properties";

    /// <summary>
    /// Styles part.
    /// </summary>
    public const string Styles = Office + "styles";

    /// <summary>
    /// Theme part.
    /// </summary>
    public const string Theme = Office + "theme";

    /// <summary>
    /// An image.
    /// </summary>
    public const string Image = Office + "image";

    /// <summary>
    /// A hyperlink target.
    /// </summary>
    public const string Hyperlink = Office + "hyperlink";

    /// <summary>
    /// A DrawingML drawing layer.
    /// </summary>
    public const string Drawing = Office + "drawing";

    /// <summary>
    /// A chart.
    /// </summary>
    public const string Chart = Office + "chart";

    /// <summary>
    /// Comments.
    /// </summary>
    public const string Comments = Office + "comments";

    /// <summary>
    /// A legacy VML drawing.
    /// </summary>
    public const string VmlDrawing = Office + "vmlDrawing";

    /// <summary>
    /// A worksheet.
    /// </summary>
    public const string Worksheet = Office + "worksheet";

    /// <summary>
    /// The shared string table of a workbook.
    /// </summary>
    public const string SharedStrings = Office + "sharedStrings";
}

/// <summary>
/// Content types shared by the Office formats.
/// </summary>
internal static class ContentTypes
{
    /// <summary>
    /// Relationship parts.
    /// </summary>
    public const string Relationships = "application/vnd.openxmlformats-package.relationships+xml";

    /// <summary>
    /// Plain XML.
    /// </summary>
    public const string Xml = "application/xml";

    /// <summary>
    /// PNG images.
    /// </summary>
    public const string Png = "image/png";

    /// <summary>
    /// Legacy VML drawings.
    /// </summary>
    public const string Vml = "application/vnd.openxmlformats-officedocument.vmlDrawing";

    /// <summary>
    /// Core properties.
    /// </summary>
    public const string CoreProperties = "application/vnd.openxmlformats-package.core-properties+xml";

    /// <summary>
    /// Extended properties.
    /// </summary>
    public const string ExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

    /// <summary>
    /// A DrawingML drawing layer.
    /// </summary>
    public const string Drawing = "application/vnd.openxmlformats-officedocument.drawing+xml";

    /// <summary>
    /// A chart.
    /// </summary>
    public const string Chart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
}
