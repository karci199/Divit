#nullable enable
namespace Divit.Packaging;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

/// <summary>
/// Writes one part's XML.
/// </summary>
/// <param name="w">Writer positioned after the XML declaration.</param>
internal delegate void XmlPartWriter(XmlWriter w);

/// <summary>
/// An Office Open XML package being written (ECMA-376 part 2, OPC): the zip, its parts, their content types. Shared
/// by every format: an xlsx, docx or pptx file is such a package with different parts.
/// </summary>
/// <remarks>
/// <para><see cref="Save"/> writes next to the target as <c>.tmp</c> and moves the file into place only when it is
/// complete, so an interrupted save never leaves a half-written document and never destroys the previous one.</para>
/// <para>Every entry carries the same time (the document's creation time when given), so identical input gives
/// identical bytes. Content types are listed in the order the caller adds them; the list is the last part.</para>
/// </remarks>
internal sealed class OpcPackage
{
    private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>
    /// Writer settings, shared: an XmlWriter only reads them. UTF-8 without byte order mark, as Office writes.
    /// </summary>
    private static readonly XmlWriterSettings WithDeclaration = new XmlWriterSettings
    {
        Encoding = new UTF8Encoding(false),
        Indent = false,
        CloseOutput = true,
    };

    private static readonly XmlWriterSettings WithoutDeclaration = new XmlWriterSettings
    {
        Encoding = new UTF8Encoding(false),
        Indent = false,
        CloseOutput = true,
        OmitXmlDeclaration = true,
    };

    private readonly ZipArchive _zip;
    private readonly DateTimeOffset? _time;
    private readonly List<(string Extension, string Type)> _defaults = new List<(string, string)>();
    private readonly List<(string Part, string Type)> _overrides = new List<(string, string)>();

    private OpcPackage(ZipArchive zip, DateTime? time)
    {
        _zip = zip;
        _time = EntryTime(time);
    }

    /// <summary>
    /// Writes a package to a file; an existing file is replaced only once the new one is complete.
    /// </summary>
    /// <param name="path">Target file.</param>
    /// <param name="time">Time stamped on every entry; the clock time when empty.</param>
    /// <param name="write">Adds the parts.</param>
    public static void Save(string path, DateTime? time, Action<OpcPackage> write)
    {
        var temp = path + ".tmp";
        try
        {
            using (var file = new FileStream(temp, FileMode.Create, FileAccess.Write, FileShare.None, 65536))
            {
                Write(file, time, write);
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            File.Move(temp, path);
        }
        finally
        {
            if (File.Exists(temp))
            {
                File.Delete(temp);
            }
        }
    }

    /// <summary>
    /// Writes a package into a stream, which stays open.
    /// </summary>
    /// <param name="stream">Writable stream.</param>
    /// <param name="time">Time stamped on every entry; the clock time when empty.</param>
    /// <param name="write">Adds the parts.</param>
    public static void Write(Stream stream, DateTime? time, Action<OpcPackage> write)
    {
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            var package = new OpcPackage(zip, time);
            write(package);
            package.XmlPart("[Content_Types].xml", package.WriteContentTypes);
        }
    }

    /// <summary>
    /// Registers a content type for every part with a file extension.
    /// </summary>
    public void Default(string extension, string type) => _defaults.Add((extension, type));

    /// <summary>
    /// Registers the content type of one part; the name starts with <c>/</c>.
    /// </summary>
    public void Override(string part, string type) => _overrides.Add((part, type));

    /// <summary>
    /// Adds an XML part. <paramref name="declaration"/> is off only for parts Office itself writes without one, such
    /// as legacy VML drawings.
    /// </summary>
    /// <param name="name">Name inside the zip, without leading <c>/</c>.</param>
    /// <param name="write">Writes the content.</param>
    /// <param name="declaration">Whether the part starts with an XML declaration (standalone).</param>
    public void XmlPart(string name, XmlPartWriter write, bool declaration = true)
    {
        var entry = Entry(name, CompressionLevel.Fastest);
        using (var w = XmlWriter.Create(entry.Open(), declaration ? WithDeclaration : WithoutDeclaration))
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
    /// Adds a relationship part, unless the list is empty.
    /// </summary>
    /// <param name="name">Name of the <c>.rels</c> part, for example <c>_rels/.rels</c>.</param>
    /// <param name="relationships">The relationships.</param>
    public void Relationships(string name, RelationshipList relationships)
    {
        if (relationships.Count > 0)
        {
            XmlPart(name, relationships.Write);
        }
    }

    /// <summary>
    /// Adds a binary part as it is. Compressed media (PNG, JPEG) are stored: deflating them again costs time and
    /// saves nothing.
    /// </summary>
    /// <param name="name">Name inside the zip.</param>
    /// <param name="data">The bytes.</param>
    /// <param name="compress">Whether to deflate the bytes.</param>
    public void BinaryPart(string name, byte[] data, bool compress)
    {
        var entry = Entry(name, compress ? CompressionLevel.Fastest : CompressionLevel.NoCompression);
        using (var stream = entry.Open())
        {
            stream.Write(data, 0, data.Length);
        }
    }

    /// <summary>
    /// Adds <c>docProps/core.xml</c> and <c>docProps/app.xml</c> with their content types.
    /// </summary>
    /// <param name="properties">The properties.</param>
    public void DocumentProperties(DocumentProperties properties)
    {
        XmlPart("docProps/core.xml", properties.WriteCore);
        XmlPart("docProps/app.xml", properties.WriteApp);
    }

    /// <summary>
    /// Zip times count in local two-second steps from 1980 to 2107; the time is stored as it is, without zone.
    /// </summary>
    private static DateTimeOffset? EntryTime(DateTime? time)
    {
        if (time is DateTime created)
        {
            var utc = created.ToUniversalTime();
            if (utc.Year >= 1980 && utc.Year <= 2107)
            {
                return new DateTimeOffset(DateTime.SpecifyKind(utc, DateTimeKind.Unspecified), TimeSpan.Zero);
            }
        }

        return null;
    }

    private ZipArchiveEntry Entry(string name, CompressionLevel level)
    {
        var entry = _zip.CreateEntry(name, level);
        if (_time != null)
        {
            entry.LastWriteTime = _time.Value;
        }

        return entry;
    }

    private void WriteContentTypes(XmlWriter w)
    {
        w.WriteStartElement("Types", ContentTypesNs);
        foreach (var item in _defaults)
        {
            w.WriteStartElement("Default", ContentTypesNs);
            w.WriteAttributeString("Extension", item.Extension);
            w.WriteAttributeString("ContentType", item.Type);
            w.WriteEndElement();
        }

        foreach (var item in _overrides)
        {
            w.WriteStartElement("Override", ContentTypesNs);
            w.WriteAttributeString("PartName", item.Part);
            w.WriteAttributeString("ContentType", item.Type);
            w.WriteEndElement();
        }

        w.WriteEndElement();
    }
}
