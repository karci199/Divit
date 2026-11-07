#nullable enable
namespace Divit.Packaging;

using System;
using System.Globalization;
using System.Xml;

/// <summary>
/// The document properties every Office file carries: core (title, author, dates) and extended (application).
/// </summary>
internal sealed class DocumentProperties
{
    private const string Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string Dc = "http://purl.org/dc/elements/1.1/";
    private const string Dcterms = "http://purl.org/dc/terms/";
    private const string Xsi = "http://www.w3.org/2001/XMLSchema-instance";
    private const string Extended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    public DocumentProperties(string application)
    {
        Application = application;
    }

    public string Application { get; }

    public string? Title { get; set; }

    public string? Subject { get; set; }

    public string? Keywords { get; set; }

    public string? Description { get; set; }

    /// <summary>
    /// Author; the application name when empty.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Creation and modification time; the time of writing when empty.
    /// </summary>
    public DateTime? Created { get; set; }

    public void WriteCore(XmlWriter w)
    {
        var now = (Created ?? DateTime.UtcNow).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        w.WriteStartElement("cp", "coreProperties", Cp);
        w.WriteAttributeString("xmlns", "dc", null, Dc);
        w.WriteAttributeString("xmlns", "dcterms", null, Dcterms);
        w.WriteAttributeString("xmlns", "xsi", null, Xsi);
        if (Title != null) { w.WriteElementString("dc", "title", Dc, OfficeText.Clean(Title)); }
        if (Subject != null) { w.WriteElementString("dc", "subject", Dc, OfficeText.Clean(Subject)); }
        w.WriteElementString("dc", "creator", Dc, OfficeText.Clean(Author ?? Application));
        if (Keywords != null) { w.WriteElementString("cp", "keywords", Cp, OfficeText.Clean(Keywords)); }
        if (Description != null) { w.WriteElementString("dc", "description", Dc, OfficeText.Clean(Description)); }
        foreach (var name in new[] { "created", "modified" })
        {
            w.WriteStartElement("dcterms", name, Dcterms);
            w.WriteAttributeString("xsi", "type", Xsi, "dcterms:W3CDTF");
            w.WriteString(now);
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }

    public void WriteApp(XmlWriter w)
    {
        w.WriteStartElement("Properties", Extended);
        w.WriteElementString("Application", Extended, OfficeText.Clean(Application));
        w.WriteEndElement();
    }
}
