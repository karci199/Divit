namespace Divit.Tests;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

/// <summary>
/// Shared helpers: temporary files, schema validation with the Open XML SDK and a gallery of sample files.
/// </summary>
internal static class XlsxTestKit
{
    /// <summary>
    /// A fixed creation time, so the same input produces the same bytes.
    /// </summary>
    internal static readonly DateTime FixedTime = new DateTime(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);

    /// <summary>
    /// A new empty folder under the system temporary folder.
    /// </summary>
    internal static string NewFolder()
    {
        var folder = Path.Combine(Path.GetTempPath(), "Divit.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    /// <summary>
    /// Schema errors of a spreadsheet as readable lines; empty when the file is valid.
    /// </summary>
    internal static List<string> Validate(string path)
    {
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            return validator.Validate(document)
                .Select(e => e.Part?.Uri + " " + e.Path?.XPath + ": " + e.Description)
                .ToList();
        }
    }

    /// <summary>
    /// Copies a file into <c>artifacts/gallery</c> at the repository root, so it can be opened and checked by eye.
    /// </summary>
    internal static void Keep(string path, string name)
    {
        var folder = AppContext.BaseDirectory;
        while (folder != null && !File.Exists(Path.Combine(folder, "Divit.sln")))
        {
            folder = Path.GetDirectoryName(folder);
        }

        if (folder == null)
        {
            return;
        }

        var gallery = Path.Combine(folder, "artifacts", "gallery");
        Directory.CreateDirectory(gallery);
        File.Copy(path, Path.Combine(gallery, name), true);
    }
}
