namespace Divit.Tests;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// Schema validation of Word documents with the Open XML SDK; the gallery and temporary folders are shared with the
/// spreadsheet tests (<see cref="XlsxTestKit"/>).
/// </summary>
internal static class DocxTestKit
{
    /// <summary>
    /// Schema errors of a document as readable lines; empty when the file is valid.
    /// </summary>
    internal static List<string> Validate(string path)
    {
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            return validator.Validate(document)
                .Select(e => e.Part?.Uri + " " + e.Path?.XPath + ": " + e.Description)
                .ToList();
        }
    }
}
