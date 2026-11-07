namespace Divit.Tests;

using System.IO;
using System.Linq;
using Xunit;
using Reader = UglyToad.PdfPig.PdfDocument;

/// <summary>
/// Word documents laid out as PDF: the demos of the Word tests saved as PDF, read back and kept for a look.
/// </summary>
public class DocxPdfTests
{
    private static string Save(Divit.Docx.DocxDocument doc, string name)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        doc.SavePdf(path);
        Assert.Empty(PdfTestKit.Validate(path));
        XlsxTestKit.Keep(path, name);
        return path;
    }

    [Fact]
    public void TheProcedureIsLaidOut()
    {
        var path = Save(DocxProcedureTests.Procedure(), "acceptance-test-procedure.pdf");
        using (var reader = Reader.Open(path))
        {
            Assert.True(reader.NumberOfPages >= 8, reader.NumberOfPages + " pages");
            var text = string.Concat(reader.GetPages().Select(p => p.Text));
            Assert.Contains("SH2000", text);
            Assert.NotNull(reader.Information.Title);
        }
    }

    [Fact]
    public void TheFeaturesAreLaidOut()
    {
        Save(DocxFeatureTests.Showcase(), "word-features.pdf");
        Save(DocxTechnicalTests.Showcase(), "word-equations-sources.pdf");
    }
}
