namespace Divit.Tests;

using Divit.Docx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

/// <summary>
/// Equations, tracked changes, sources, citations and the bibliography.
/// </summary>
public class DocxTechnicalTests
{
    private static readonly XNamespace B = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

    private static string Save(DocxDocument doc, string name, bool keep = false)
    {
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        doc.Save(path);
        Assert.Empty(DocxTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    /// <summary>
    /// The link budget of an antenna test: equations with numbers, sources cited, a revised sentence.
    /// </summary>
    internal static DocxDocument Showcase(DocxBibliographyStyle style = DocxBibliographyStyle.Ieee)
    {
        var doc = new DocxDocument { Created = XlsxTestKit.FixedTime, NumberHeadings = true, Author = "Test team", BibliographyStyle = style };
        doc.Look.BodyFont = "Trebuchet MS";
        var balanis = doc.AddSource("Balanis2016", DocxSourceKind.Book, "Antenna Theory: Analysis and Design")
            .AddAuthor("Balanis", "Constantine A.");
        balanis.Year = "2016";
        balanis.Edition = "4th";
        balanis.Publisher = "Wiley";
        balanis.City = "Hoboken";
        var ieee = doc.AddSource("IEEE149", DocxSourceKind.Report, "IEEE Recommended Practice for Antenna Measurements");
        ieee.Corporate = "IEEE";
        ieee.Year = "2021";
        ieee.Publisher = "IEEE Standards Association";
        var paper = doc.AddSource("Kummer1978", DocxSourceKind.JournalArticle, "Antenna measurements — 1978")
            .AddAuthor("Kummer", "W. H.").AddAuthor("Gillespie", "E. S.");
        paper.Journal = "Proceedings of the IEEE";
        paper.Volume = "66";
        paper.Issue = "4";
        paper.Pages = "483–507";
        paper.Year = "1978";

        doc.AddHeading("Link budget", 1);
        doc.AddParagraph().Add("The received power follows from the Friis equation ").AddCitation(balanis, "94").Add(":");
        var friis = doc.AddEquation(
            DocxMath.Row(
                DocxMath.Sub("P", "r"), "=",
                DocxMath.Sub("P", "t"), "+", DocxMath.Sub("G", "t"), "+", DocxMath.Sub("G", "r"), "+",
                DocxMath.Plain("20"),
                DocxMath.Function(DocxMath.Sub(DocxMath.Plain("log"), DocxMath.Plain("10")),
                    DocxMath.Brackets(DocxMath.Fraction("λ", DocxMath.Row(DocxMath.Plain("4"), "πR"))))),
            numbered: true);
        doc.AddParagraph().Add("The far field starts at the distance ").AddCitation(ieee).Add(":");
        var far = doc.AddEquation(DocxMath.Row("R", "≥", DocxMath.Fraction(DocxMath.Row(DocxMath.Plain("2"), DocxMath.Sup("D", "2")), "λ")), numbered: true);
        doc.AddParagraph().Add("Equations ").AddReference(friis, DocxReferenceKind.Label).Add(" and ").AddReference(far, DocxReferenceKind.Label)
            .Add(" hold in the chamber ").AddCitation(paper).Add(". A gain of ").AddEquation(DocxMath.Row("G", "=", DocxMath.Plain("8 dBi"))).Add(" is expected.");
        doc.AddEquation(DocxMath.Row(
            "η", "=", DocxMath.Fraction(DocxMath.Sub("P", DocxMath.Plain("rad")), DocxMath.Sub("P", DocxMath.Plain("in"))), ",", " ",
            DocxMath.Root(DocxMath.Row("x", "+", DocxMath.Plain("1"))), ",", " ",
            DocxMath.Root("x", DocxMath.Plain("3")), ",", " ",
            DocxMath.Sum(DocxMath.Sub("a", "i"), DocxMath.Row("i", "=", DocxMath.Plain("1")), "N"), ",", " ",
            DocxMath.Integral(DocxMath.Row("f", DocxMath.Brackets("t"), DocxMath.Plain("d"), "t"), DocxMath.Plain("0"), "T"), ",", " ",
            DocxMath.Brackets(DocxMath.SubSup("S", DocxMath.Plain("11"), DocxMath.Plain("2")), "|", "|"), ",", " ",
            DocxMath.Integral("g")));

        doc.AddHeading("Revision", 1);
        doc.AddParagraph().Add("The antenna is measured at ")
            .AddDeletion("8.0", "Reviewer").AddInsertion("8.2", "Reviewer", new DateTime(2026, 2, 3, 10, 0, 0, DateTimeKind.Utc))
            .Add(" GHz").AddInsertion(" and 8.4 GHz").Add(".");
        doc.AddBibliography();
        return doc;
    }

    [Fact]
    public void TheShowcaseIsValid()
    {
        Save(Showcase(), "word-equations-sources.docx", keep: true);
        Save(Showcase(DocxBibliographyStyle.Apa), "word-sources-apa.docx");
    }

    [Fact]
    public void EquationsAreOfficeMath()
    {
        var path = Save(Showcase(), "math.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            var maths = body.Descendants<M.OfficeMath>().ToList();
            Assert.Equal(4, maths.Count);
            Assert.Equal(3, body.Descendants<M.Paragraph>().Count());
            Assert.Contains("xmlns:m=", document.MainDocumentPart.Document.OuterXml.Substring(0, 400));

            var friis = maths[0];
            Assert.Equal(4, friis.Descendants<M.Subscript>().Count(s => s.Parent is M.OfficeMath));
            var function = friis.Descendants<M.MathFunction>().Single();
            Assert.Equal("log10", function.FunctionName!.InnerText);
            Assert.Equal("p", function.FunctionName.Descendants<M.Style>().First().Val!.InnerText);
            var fraction = function.Descendants<M.Fraction>().Single();
            Assert.Equal("λ", fraction.Numerator!.InnerText);
            Assert.Equal("4πR", fraction.Denominator!.InnerText);
            Assert.Null(function.Descendants<M.Delimiter>().Single().DelimiterProperties);

            var last = maths[3];
            var roots = last.Descendants<M.Radical>().ToList();
            Assert.Equal("1", roots[0].RadicalProperties!.HideDegree!.Val!.InnerText);
            Assert.Equal(string.Empty, roots[0].Degree!.InnerText);
            Assert.Null(roots[1].RadicalProperties);
            Assert.Equal("3", roots[1].Degree!.InnerText);
            var operators = last.Descendants<M.Nary>().ToList();
            Assert.Equal("∑", operators[0].NaryProperties!.AccentChar!.Val!.Value);
            Assert.Equal("undOvr", operators[0].NaryProperties!.LimitLocation!.Val!.InnerText);
            Assert.Equal("i=1", operators[0].SubArgument!.InnerText);
            Assert.Equal("subSup", operators[1].NaryProperties!.LimitLocation!.Val!.InnerText);
            Assert.Null(operators[1].NaryProperties!.HideSubArgument);
            Assert.Equal("1", operators[2].NaryProperties!.HideSubArgument!.Val!.InnerText);
            Assert.Equal("1", operators[2].NaryProperties!.HideSuperArgument!.Val!.InnerText);
            var bars = last.Descendants<M.Delimiter>().Single(d => d.DelimiterProperties != null);
            Assert.Equal("|", bars.DelimiterProperties!.BeginChar!.Val!.Value);
            Assert.Single(bars.Descendants<M.SubSuperscript>());

            // Inline in a sentence, without a paragraph of its own.
            Assert.IsType<Paragraph>(maths[2].Parent);
            Assert.Contains("G=8 dBi", maths[2].InnerText);
        }
    }

    [Fact]
    public void NumberedEquationsAreCaptions()
    {
        var path = Save(Showcase(), "math-numbers.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            var numbered = body.Descendants<Paragraph>().Where(p => p.Descendants<SimpleField>().Any(f => f.Instruction!.Value!.Contains("SEQ Equation"))).ToList();
            Assert.Equal(2, numbered.Count);
            Assert.Equal(new[] { "(1)", "(2)" }, numbered.Select(p => p.InnerText).ToArray());

            // A row without lines: the equation at full size in the middle, its number at the right.
            var row = (TableRow)numbered[0].Parent!.Parent!;
            var cells = row.Elements<TableCell>().ToList();
            Assert.Equal(3, cells.Count);
            Assert.Equal("DivitPlain", ((Table)row.Parent!).GetFirstChild<TableProperties>()!.TableStyle!.Val!.Value);
            Assert.Single(cells[1].Descendants<M.Paragraph>());
            Assert.Same(cells[2], numbered[0].Parent);
            Assert.Equal(JustificationValues.Right, numbered[0].ParagraphProperties!.Justification!.Val!.Value);

            var references = body.Descendants<SimpleField>().Where(f => f.Instruction!.Value!.StartsWith(" REF")).ToList();
            Assert.Equal(new[] { "(1)", "(2)" }, references.Select(r => r.InnerText).ToArray());
            var name = references[0].Instruction!.Value!.Split(' ')[2];
            var start = numbered[0].Elements<BookmarkStart>().Single(b => b.Name == name);
            var inside = numbered[0].ChildElements.SkipWhile(e => e != start).Skip(1).TakeWhile(e => !(e is BookmarkEnd)).ToList();
            Assert.Equal("(1)", string.Concat(inside.Select(e => e.InnerText)));
        }
    }

    [Fact]
    public void ChangesAreTracked()
    {
        var path = Save(Showcase(), "revisions.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            var deleted = body.Descendants<DeletedRun>().Single();
            Assert.Equal("8.0", deleted.Descendants<DeletedText>().Single().Text);
            Assert.Equal("Reviewer", deleted.Author!.Value);
            Assert.Equal(XlsxTestKit.FixedTime, deleted.Date!.Value.ToUniversalTime());
            var inserted = body.Descendants<InsertedRun>().ToList();
            Assert.Equal(new[] { "8.2", " and 8.4 GHz" }, inserted.Select(i => i.InnerText).ToArray());
            Assert.Equal(new DateTime(2026, 2, 3, 10, 0, 0, DateTimeKind.Utc), inserted[0].Date!.Value.ToUniversalTime());
            Assert.Equal("Test team", inserted[1].Author!.Value);

            var ids = body.Descendants<BookmarkStart>().Select(b => b.Id!.Value)
                .Concat(new[] { deleted.Id!.Value }).Concat(inserted.Select(i => i.Id!.Value)).ToList();
            Assert.Equal(ids.Count, ids.Distinct().Count());
            var revised = (Paragraph)deleted.Parent!;
            Assert.DoesNotContain("8.0", string.Concat(revised.Descendants<Text>().Select(t => t.Text)));
        }

        // Deleted text is gone from the paragraph's text, which contents and references show.
        var heading = new DocxDocument().AddHeading("Old");
        heading.AddDeletion(" title").AddInsertion(" name");
        Assert.Equal("Old name", heading.Text);

        var plain = new DocxDocument();
        plain.AddParagraph().AddInsertion("x");
        var undated = Save(plain, "undated.docx");
        using (var document = WordprocessingDocument.Open(undated, false))
        {
            var inserted = document.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>().Single();
            Assert.Null(inserted.Date);
            Assert.Equal("Divit", inserted.Author!.Value);
        }
    }

    [Fact]
    public void SourcesLiveInACustomXmlPart()
    {
        var path = Save(Showcase(), "sources.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var item = main.CustomXmlParts.Single();
            Assert.Equal(BibliographyNsUri, item.CustomXmlPropertiesPart!.DataStoreItem!.SchemaReferences!.Elements<DocumentFormat.OpenXml.CustomXmlDataProperties.SchemaReference>().Single().Uri!.Value);
            XDocument xml;
            using (var stream = item.GetStream())
            {
                xml = XDocument.Load(stream);
            }

            Assert.Equal("IEEE", xml.Root!.Attribute("StyleName")!.Value);
            var sources = xml.Root.Elements(B + "Source").ToList();
            Assert.Equal(new[] { "Balanis2016", "IEEE149", "Kummer1978" }, sources.Select(s => s.Element(B + "Tag")!.Value).ToArray());
            Assert.Equal(new[] { "Book", "Report", "JournalArticle" }, sources.Select(s => s.Element(B + "SourceType")!.Value).ToArray());
            Assert.Equal("Balanis", sources[0].Descendants(B + "Last").Single().Value);
            Assert.Equal("IEEE", sources[1].Descendants(B + "Corporate").Single().Value);
            Assert.Equal(2, sources[2].Descendants(B + "Person").Count());
            Assert.Equal("483–507", sources[2].Element(B + "Pages")!.Value);
            Assert.Equal(new[] { "1", "2", "3" }, sources.Select(s => s.Element(B + "RefOrder")!.Value).ToArray());

            var citations = main.Document!.Body!.Descendants<SimpleField>().Where(f => f.Instruction!.Value!.Contains("CITATION")).ToList();
            Assert.Equal(new[] { " CITATION Balanis2016 \\l 1033 \\p 94 ", " CITATION IEEE149 \\l 1033 ", " CITATION Kummer1978 \\l 1033 " }, citations.Select(c => c.Instruction!.Value).ToArray());
            Assert.Equal(new[] { "[1, p. 94]", "[2]", "[3]" }, citations.Select(c => c.InnerText).ToArray());

            var bibliography = main.Document.Body.Elements<Paragraph>().Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Bibliography").ToList();
            Assert.Equal(3, bibliography.Count);
            Assert.Contains(" BIBLIOGRAPHY ", bibliography[0].InnerText);
            Assert.EndsWith("[1]C. A. Balanis, “Antenna Theory: Analysis and Design,” 4th ed., Hoboken, Wiley, 2016.", bibliography[0].InnerText);
            Assert.EndsWith("[3]W. H. Kummer and E. S. Gillespie, “Antenna measurements — 1978,” Proceedings of the IEEE, vol. 66, no. 4, pp. 483–507, 1978.", bibliography[2].InnerText);
            Assert.Equal("true", main.DocumentSettingsPart!.Settings!.GetFirstChild<UpdateFieldsOnOpen>()!.Val!.InnerText);
        }
    }

    [Fact]
    public void ApaCitesAuthorAndYear()
    {
        var doc = Showcase(DocxBibliographyStyle.Apa);
        doc.Language = "tr-TR";
        var path = Save(doc, "apa.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            using (var stream = main.CustomXmlParts.Single().GetStream())
            {
                Assert.Equal("APA", XDocument.Load(stream).Root!.Attribute("StyleName")!.Value);
            }

            var citations = main.Document!.Body!.Descendants<SimpleField>().Where(f => f.Instruction!.Value!.Contains("CITATION")).ToList();
            Assert.Equal(new[] { "(Balanis, 2016, p. 94)", "(IEEE, 2021)", "(Kummer & Gillespie, 1978)" }, citations.Select(c => c.InnerText).ToArray());
            Assert.Contains("\\l 1055", citations[0].Instruction!.Value);
            var first = main.Document.Body.Elements<Paragraph>().First(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Bibliography");
            Assert.EndsWith("Balanis, C. A. (2016). Antenna Theory: Analysis and Design. Hoboken: Wiley.", first.InnerText);
        }
    }

    [Fact]
    public void SourcesAreChecked()
    {
        var doc = new DocxDocument();
        doc.AddSource("A", DocxSourceKind.Book, "One");
        Assert.Throws<ArgumentException>(() => doc.AddSource("a", DocxSourceKind.Book, "Same tag"));
        Assert.Throws<ArgumentException>(() => doc.AddSource("Two words", DocxSourceKind.Book, "Space"));
        Assert.Throws<ArgumentException>(() => doc.AddSource(string.Empty, DocxSourceKind.Book, "Empty"));

        var foreign = new DocxDocument().AddSource("B", DocxSourceKind.Misc, "Elsewhere");
        doc.AddParagraph().AddCitation(foreign);
        Assert.Throws<InvalidOperationException>(() => doc.Save(Path.Combine(XlsxTestKit.NewFolder(), "foreign.docx")));
    }

    [Fact]
    public void WithoutSourcesThereIsNoCustomXml()
    {
        var doc = new DocxDocument { Created = XlsxTestKit.FixedTime };
        doc.AddEquation("x");
        var path = Save(doc, "no-sources.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            Assert.Empty(document.MainDocumentPart!.CustomXmlParts);
            Assert.Null(document.MainDocumentPart.DocumentSettingsPart!.Settings!.GetFirstChild<UpdateFieldsOnOpen>());
        }
    }

    [Fact]
    public void TheShowcaseIsReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        Showcase().Save(Path.Combine(folder, "a.docx"));
        Showcase().Save(Path.Combine(folder, "b.docx"));
        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.docx")), File.ReadAllBytes(Path.Combine(folder, "b.docx")));
    }

    private const string BibliographyNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
}
