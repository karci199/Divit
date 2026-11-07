namespace Divit.Tests;

using Divit.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using Xunit;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

/// <summary>
/// Lists, links, cross-references, watermarks, check boxes, footnotes, protection and change tracking.
/// </summary>
public class DocxFeatureTests
{
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

    private static DocxDocument NewDocument(bool numberHeadings = false) =>
        new DocxDocument { Created = XlsxTestKit.FixedTime, NumberHeadings = numberHeadings };

    /// <summary>
    /// A page that uses every feature of this file.
    /// </summary>
    private static DocxDocument Showcase()
    {
        var doc = NewDocument(true);
        doc.Look.BodyFont = "Trebuchet MS";
        doc.Title = "Word features";
        doc.Section.Watermark = "DRAFT";
        doc.Section.Header.AddParagraph("Divit · Word features").Add("  ").AddLink("divit.example", "https://divit.example");
        doc.Section.Footer.AddParagraph().Add("Page ").AddField(DocxField.Page);

        var lists = doc.AddHeading("Lists", 1);
        doc.AddParagraph("Bullets, three levels deep:");
        var bullets = doc.AddList();
        bullets.AddItem("Network analyser");
        bullets.AddItem("Port 1", 1);
        bullets.AddItem("Calibration kit", 2);
        bullets.AddItem("Absorber");
        doc.AddParagraph("Numbered steps, each list starting at 1:");
        var steps = doc.AddList(DocxListKind.Numbered);
        steps.AddItem("Switch on the analyser.");
        steps.AddItem("Let it warm up for 30 minutes.", 1);
        steps.AddItem("Calibrate port 1.");
        var again = doc.AddList(DocxListKind.Numbered);
        again.AddItem("A second list starts again at 1.");

        doc.AddHeading("Figures and references", 1);
        doc.AddParagraph().AddImage(TestImages.SetUp(), 12, description: "Test set-up");
        var figure = doc.AddCaption(DocxCaptionKind.Figure, "Test set-up");
        var refs = doc.AddHeading("Cross-references", 2);
        doc.AddParagraph().Add("The set-up is in ").AddReference(figure, DocxReferenceKind.Label).Add(" on page ")
            .AddReference(figure, DocxReferenceKind.Page).Add("; lists are explained in chapter ")
            .AddReference(lists, DocxReferenceKind.Number).Add(" (").AddReference(lists).Add(").");
        doc.AddParagraph().Add("Links: ").AddLink("the project site", "https://divit.example").Add(", ")
            .AddLink("e-mail", "mailto:test@divit.example").Add(", ").AddLink("back to the lists", lists).Add(".");
        doc.AddParagraph().Add("A footnote follows this sentence.").AddFootnote("Footnotes sit at the foot of the page.")
            .Add(" And another one.").AddFootnote("Numbered by Word: 1, 2, 3.");

        doc.AddHeading("Record", 1);
        var table = doc.AddTable(3).SetColumnWidths(8, 3, 3);
        table.AddHeaderRow("Check", "Pass", "Fail");
        foreach (var (check, pass) in new[] { ("Return loss ≤ −10 dB", true), ("Gain ≥ 8 dBi", false) })
        {
            var row = table.AddRow(check);
            row.Cells[1].SetText(string.Empty).AddCheckBox(pass).Alignment = DocxAlignment.Center;
            row.Cells[2].SetText(string.Empty).AddCheckBox().Alignment = DocxAlignment.Center;
        }

        doc.AddParagraph().Add("Reviewed ").AddCheckBox(false, new DocxFont { Size = 14, Color = "#1F4E79" });
        _ = refs;
        return doc;
    }

    [Fact]
    public void TheShowcaseIsValid()
    {
        var doc = Showcase();
        doc.Protect(DocxProtection.Forms);
        Save(doc, "word-features.docx", keep: true);
    }

    [Fact]
    public void ListsGetANumberingEach()
    {
        var path = Save(Showcase(), "lists.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            Assert.Equal(new[] { 0, 1, 2 }, numbering.Elements<AbstractNum>().Select(a => a.AbstractNumberId!.Value).ToArray());
            var nums = numbering.Elements<NumberingInstance>().ToList();
            Assert.Equal(new[] { 1, 2, 3, 4 }, nums.Select(n => n.NumberID!.Value).ToArray());
            Assert.Null(nums[1].GetFirstChild<LevelOverride>());
            Assert.Equal(1, nums[2].GetFirstChild<LevelOverride>()!.StartOverrideNumberingValue!.Val!.Value);
            Assert.Equal(1, nums[3].GetFirstChild<LevelOverride>()!.StartOverrideNumberingValue!.Val!.Value);

            var bullet = numbering.Elements<AbstractNum>().ElementAt(1).Elements<Level>().ToList();
            Assert.Equal(9, bullet.Count);
            Assert.Equal("•", bullet[0].LevelText!.Val!.Value);
            Assert.Equal("–", bullet[1].LevelText!.Val!.Value);
            var numbered = numbering.Elements<AbstractNum>().ElementAt(2).Elements<Level>().ToList();
            Assert.Equal("%2)", numbered[1].LevelText!.Val!.Value);
            Assert.Equal(NumberFormatValues.LowerLetter, numbered[1].NumberingFormat!.Val!.Value);

            var items = main.Document!.Body!.Descendants<Paragraph>().Where(p => p.ParagraphProperties?.NumberingProperties != null).ToList();
            Assert.Equal(8, items.Count);
            Assert.All(items, p => Assert.Equal(DocxStyle.ListParagraph, p.ParagraphProperties!.ParagraphStyleId!.Val!.Value));
            Assert.Equal(new[] { 0, 1, 2, 0, 0, 1, 0, 0 }, items.Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value).ToArray());
            Assert.Equal(new[] { 2, 2, 2, 2, 3, 3, 3, 4 }, items.Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).ToArray());
            Assert.NotNull(main.StyleDefinitionsPart!.Styles!.Elements<Style>().SingleOrDefault(s => s.StyleId == DocxStyle.ListParagraph));
        }
    }

    [Fact]
    public void ListsWorkWithoutHeadingNumbersAndInCells()
    {
        var doc = NewDocument();
        var cell = doc.AddTable(1).AddRow().Cells[0];
        cell.AddList(DocxListKind.Numbered).AddItem("In a cell", 12);
        var path = Save(doc, "list-cell.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var numbering = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            Assert.Equal(2, numbering.Elements<AbstractNum>().Single().AbstractNumberId!.Value);
            Assert.Equal(2, numbering.Elements<NumberingInstance>().Single().NumberID!.Value);
            var item = document.MainDocumentPart.Document!.Body!.Descendants<NumberingProperties>().Single();
            Assert.Equal(8, item.NumberingLevelReference!.Val!.Value);
        }
    }

    [Fact]
    public void LinksPointOutOfTheDocumentOrToBookmarks()
    {
        var path = Save(Showcase(), "links.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var links = main.Document!.Body!.Descendants<Hyperlink>().ToList();
            Assert.Equal(3, links.Count);
            var site = main.HyperlinkRelationships.Single(r => r.Id == links[0].Id!.Value);
            Assert.Equal(new Uri("https://divit.example"), site.Uri);
            Assert.True(site.IsExternal);
            Assert.Equal("mailto:test@divit.example", main.HyperlinkRelationships.Single(r => r.Id == links[1].Id!.Value).Uri.OriginalString);
            Assert.Null(links[2].Id);
            var anchor = links[2].Anchor!.Value;
            var start = main.Document.Body.Descendants<BookmarkStart>().Single(b => b.Name == anchor);
            Assert.Equal("Lists", ((Paragraph)start.Parent!).InnerText);
            Assert.All(links, l => Assert.Equal("Hyperlink", l.Descendants<RunStyle>().Single().Val!.Value));

            // A link in the header is a relationship of the header part.
            var header = main.HeaderParts.Single();
            var headerLink = header.Header!.Descendants<Hyperlink>().Single();
            Assert.Equal(new Uri("https://divit.example"), header.HyperlinkRelationships.Single(r => r.Id == headerLink.Id!.Value).Uri);
            Assert.NotNull(main.StyleDefinitionsPart!.Styles!.Elements<Style>().SingleOrDefault(s => s.StyleId == "Hyperlink"));
        }
    }

    [Fact]
    public void ReferencesAreFieldsOnHiddenBookmarks()
    {
        var path = Save(Showcase(), "references.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var body = main.Document!.Body!;
            var fields = body.Descendants<SimpleField>().Where(f => f.Instruction!.Value!.Contains("REF")).ToList();
            Assert.Equal(4, fields.Count);
            var label = fields[0].Instruction!.Value!;
            Assert.Matches(@"^ REF _Ref\d+ \\h $", label);
            Assert.Equal("Figure 1", fields[0].InnerText);
            Assert.StartsWith(" PAGEREF _Ref", fields[1].Instruction!.Value);
            Assert.Equal("1", fields[1].InnerText);
            Assert.EndsWith(" \\r \\h ", fields[2].Instruction!.Value);
            Assert.Equal("1", fields[2].InnerText);
            Assert.Equal("Lists", fields[3].InnerText);

            // The label bookmark holds the label and the number only; the page bookmark the whole caption.
            var labelName = label.Split(' ')[2];
            var labelStart = body.Descendants<BookmarkStart>().Single(b => b.Name == labelName);
            var caption = (Paragraph)labelStart.Parent!;
            var inside = caption.ChildElements.SkipWhile(e => e != labelStart).Skip(1)
                .TakeWhile(e => !(e is BookmarkEnd end && end.Id == labelStart.Id)).ToList();
            Assert.Equal("Figure 1", string.Concat(inside.Select(e => e.InnerText)));
            var pageName = fields[1].Instruction!.Value!.Split(' ')[2];
            var pageStart = body.Descendants<BookmarkStart>().Single(b => b.Name == pageName);
            Assert.Same(caption, pageStart.Parent);
            Assert.IsType<BookmarkEnd>(caption.LastChild);

            var ids = body.Descendants<BookmarkStart>().Select(b => b.Id!.Value).ToList();
            Assert.Equal(ids.Count, ids.Distinct().Count());
            Assert.Equal(ids.Count, body.Descendants<BookmarkEnd>().Count());
            Assert.Equal("true", main.DocumentSettingsPart!.Settings!.GetFirstChild<UpdateFieldsOnOpen>()!.Val!.InnerText);
        }
    }

    [Fact]
    public void ReferencesNeedTheirTargetInTheBody()
    {
        var doc = NewDocument();
        var lost = doc.Section.Header.AddParagraph("In the header");
        doc.AddParagraph().AddReference(lost);
        Assert.Throws<InvalidOperationException>(() => doc.Save(Path.Combine(XlsxTestKit.NewFolder(), "lost.docx")));

        var other = NewDocument().AddParagraph("Elsewhere");
        var link = NewDocument();
        link.AddParagraph().AddLink("away", other);
        Assert.Throws<InvalidOperationException>(() => link.Save(Path.Combine(XlsxTestKit.NewFolder(), "away.docx")));

        var plain = NewDocument();
        var paragraph = plain.AddParagraph("No caption");
        Assert.Throws<ArgumentException>(() => plain.AddParagraph().AddReference(paragraph, DocxReferenceKind.Label));
        Assert.Throws<ArgumentException>(() => plain.AddParagraph().AddLink("x", string.Empty));
    }

    [Fact]
    public void WatermarksGoIntoTheHeaders()
    {
        var doc = NewDocument();
        doc.AddParagraph("Cover");
        doc.Section.Watermark = "CONFIDENTIAL";
        doc.Section.WatermarkColor = "#FF0000";
        doc.AddSection().AddParagraph("Body");
        doc.Section.Header.AddParagraph("Own header, no watermark");
        var path = Save(doc, "watermark.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var headers = main.HeaderParts.Select(h => h.Header!).OrderBy(h => h.InnerText).ToList();
            Assert.Equal(2, headers.Count);
            var marked = headers.Single(h => h.OuterXml.Contains("textpath"));
            Assert.Contains("string=\"CONFIDENTIAL\"", marked.OuterXml);
            Assert.Contains("fillcolor=\"#FF0000\"", marked.OuterXml);
            Assert.Contains("rotation:315", marked.OuterXml);
            Assert.Contains("_x0000_t136", marked.OuterXml);
            Assert.DoesNotContain("textpath", headers.Single(h => h != marked).OuterXml);
            Assert.Equal(2, main.Document!.Body!.Descendants<HeaderReference>().Count());
        }
    }

    [Fact]
    public void CheckBoxesAreContentControls()
    {
        var path = Save(Showcase(), "checkboxes.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var body = document.MainDocumentPart!.Document!;
            var boxes = body.Descendants<SdtRun>().ToList();
            Assert.Equal(5, boxes.Count);
            var states = boxes.Select(b => b.SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>()!.Checked!.Val!.Value).ToArray();
            Assert.Equal(new[] { W14.OnOffValues.One, W14.OnOffValues.Zero, W14.OnOffValues.Zero, W14.OnOffValues.Zero, W14.OnOffValues.Zero }, states);
            Assert.Equal("☒", boxes[0].InnerText);
            Assert.Equal("☐", boxes[1].InnerText);
            var ids = boxes.Select(b => b.SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value).ToList();
            Assert.Equal(ids.Count, ids.Distinct().Count());
            var last = boxes[4].Descendants<RunProperties>().Single();
            Assert.Equal("MS Gothic", last.RunFonts!.Ascii!.Value);
            Assert.Equal("28", last.FontSize!.Val!.Value);
            Assert.Contains("mc:Ignorable=\"w14\"", body.OuterXml.Substring(0, 600));
        }
    }

    [Fact]
    public void FootnotesHaveTheirOwnPart()
    {
        var doc = Showcase();
        doc.Section.Footer.AddParagraph("Footer note").AddFootnote("Left out: no footnotes in a footer.");
        var path = Save(doc, "footnotes.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            var notes = main.FootnotesPart!.Footnotes!.Elements<Footnote>().ToList();
            Assert.Equal(new long[] { -1, 0, 1, 2 }, notes.Select(n => n.Id!.Value).ToArray());
            Assert.Equal(FootnoteEndnoteValues.Separator, notes[0].Type!.Value);
            Assert.Equal(FootnoteEndnoteValues.ContinuationSeparator, notes[1].Type!.Value);
            Assert.Equal(" Footnotes sit at the foot of the page.", notes[2].InnerText);
            Assert.NotNull(notes[2].Descendants<FootnoteReferenceMark>().SingleOrDefault());
            Assert.Equal(new long[] { 1, 2 }, main.Document!.Body!.Descendants<FootnoteReference>().Select(r => r.Id!.Value).ToArray());
            Assert.Empty(main.FooterParts.Single().Footer!.Descendants<FootnoteReference>());
            Assert.Equal(2, main.DocumentSettingsPart!.Settings!.GetFirstChild<FootnoteDocumentWideProperties>()!.Elements<FootnoteSpecialReference>().Count());
            var styles = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Select(s => s.StyleId!.Value).ToList();
            Assert.Contains("FootnoteText", styles);
            Assert.Contains("FootnoteReference", styles);
        }
    }

    [Theory]
    [InlineData(DocxProtection.ReadOnly, "readOnly")]
    [InlineData(DocxProtection.Forms, "forms")]
    [InlineData(DocxProtection.Comments, "comments")]
    [InlineData(DocxProtection.TrackedChanges, "trackedChanges")]
    public void ProtectionIsASetting(DocxProtection protection, string edit)
    {
        var doc = NewDocument();
        doc.AddParagraph("Locked");
        doc.Protect(protection, "secret");
        doc.TrackChanges = true;
        var path = Save(doc, "protected.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var settings = document.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            var lockElement = settings.GetFirstChild<DocumentProtection>()!;
            Assert.Equal(edit, lockElement.Edit!.InnerText);
            Assert.True(lockElement.Enforcement!.Value);
            Assert.Equal("SHA-512", lockElement.AlgorithmName!.Value);
            Assert.Equal(100000L, (long)lockElement.SpinCount!.Value);
            var salt = Convert.FromBase64String(lockElement.SaltValue!.Value!);
            Assert.Equal(Divit.Packaging.PasswordHash.Sha512("secret", salt), lockElement.HashValue!.Value);
            Assert.NotNull(settings.GetFirstChild<TrackRevisions>());
            Assert.True(settings.ChildElements.ToList().IndexOf(settings.GetFirstChild<TrackRevisions>()!) < settings.ChildElements.ToList().IndexOf(lockElement));
        }
    }

    [Fact]
    public void ProtectionWithoutPasswordHasNoHash()
    {
        var doc = NewDocument();
        doc.Protect();
        var path = Save(doc, "unlocked.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var lockElement = document.MainDocumentPart!.DocumentSettingsPart!.Settings!.GetFirstChild<DocumentProtection>()!;
            Assert.Equal("readOnly", lockElement.Edit!.InnerText);
            Assert.Null(lockElement.HashValue);
            Assert.Null(document.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<TrackRevisions>());
        }
    }

    [Fact]
    public void UnusedFeaturesLeaveNoTrace()
    {
        var doc = NewDocument();
        doc.AddParagraph("Plain");
        var path = Save(doc, "plain.docx");
        using (var document = WordprocessingDocument.Open(path, false))
        {
            var main = document.MainDocumentPart!;
            Assert.Null(main.NumberingDefinitionsPart);
            Assert.Null(main.FootnotesPart);
            Assert.Empty(main.HeaderParts);
            var styles = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Select(s => s.StyleId!.Value).ToList();
            Assert.DoesNotContain("Hyperlink", styles);
            Assert.DoesNotContain(DocxStyle.ListParagraph, styles);
            Assert.DoesNotContain("FootnoteText", styles);
            Assert.Null(main.DocumentSettingsPart!.Settings!.GetFirstChild<DocumentProtection>());
            Assert.DoesNotContain("w14", main.Document!.OuterXml);
        }
    }

    [Fact]
    public void FeaturesAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        foreach (var name in new[] { "a.docx", "b.docx" })
        {
            var doc = Showcase();
            doc.Protect(DocxProtection.Forms, "secret");
            doc.Save(Path.Combine(folder, name));
        }

        Assert.Equal(File.ReadAllBytes(Path.Combine(folder, "a.docx")), File.ReadAllBytes(Path.Combine(folder, "b.docx")));
    }
}
