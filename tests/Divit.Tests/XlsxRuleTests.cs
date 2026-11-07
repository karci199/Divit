namespace Divit.Tests;

using Divit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using Xunit;

/// <summary>
/// Formulas, data validation and conditional formatting: what is written for each kind of rule, the limits Excel
/// sets, and that sheets using them all at once stay valid.
/// </summary>
public class XlsxRuleTests
{
    private static readonly XlsxStyle Green = new XlsxStyle { Fill = "#C6EFCE", FontColor = "#006100" };
    private static readonly XlsxStyle Red = new XlsxStyle { Fill = "#FFC7CE", FontColor = "#9C0006" };
    private static readonly XlsxStyle Yellow = new XlsxStyle { Fill = "#FFEB9C", FontColor = "#9C5700" };

    private static string Save(XlsxWorkbook book, string name, bool keep = false)
    {
        book.Created ??= XlsxTestKit.FixedTime;
        var path = Path.Combine(XlsxTestKit.NewFolder(), name);
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
        if (keep)
        {
            XlsxTestKit.Keep(path, name);
        }

        return path;
    }

    private static Worksheet FirstSheet(SpreadsheetDocument document) =>
        document.WorkbookPart!.WorksheetParts.First().Worksheet;

    private static Cell CellOf(Worksheet sheet, string reference) =>
        sheet.Descendants<Cell>().Single(c => c.CellReference == reference);

    #region Formulas

    /// <summary>
    /// An order list with totals: plain formulas, a defined name, cached results of both kinds.
    /// </summary>
    private static XlsxWorkbook Orders()
    {
        var book = new XlsxWorkbook { Title = "Formulas" };
        var sheet = book.AddSheet("Orders");
        sheet.Set(1, 1, "Item");
        sheet.Set(1, 2, "Quantity");
        sheet.Set(1, 3, "Price");
        sheet.Set(1, 4, "Amount");
        sheet.Format(1, 1, 1, 4, (ref XlsxStyle s) => { s.Bold = true; s.Fill = "#DCE6F1"; });
        var items = new[] { ("Pen", 12.0, 1.5), ("Paper", 3.0, 4.25), ("Ink", 2.0, 9.9), ("Ruler", 5.0, 2.0) };
        for (var i = 0; i < items.Length; i++)
        {
            var row = i + 2;
            sheet.Set(row, 1, items[i].Item1);
            sheet.Set(row, 2, items[i].Item2);
            sheet.Set(row, 3, items[i].Item3);
            sheet.SetFormula(row, 4, "=B" + row + "*C" + row);
        }

        book.DefineName("Amounts", sheet, 2, 4, 5, 4);
        sheet.Set(7, 3, "Total");
        sheet.SetFormula(7, 4, "SUM(Amounts)", 70.3);
        sheet.Set(8, 3, "Largest");
        sheet.SetFormula(8, 4, " = MAX(D2:D5) ");
        sheet.Set(9, 3, "Status");
        sheet.SetFormula(9, 4, "IF(D7>50,\"Over budget\",\"OK\")", "Over budget");
        sheet.Set(10, 3, "Any large?");
        sheet.SetFormula(10, 4, "OR(D2:D5>20)");
        sheet.Format(2, 3, 8, 4, (ref XlsxStyle s) => s.NumberFormat = "#,##0.00");
        sheet.ColumnWidth(4, 14);
        return book;
    }

    [Fact]
    public void FormulasAreWrittenWithAndWithoutResults()
    {
        var path = Save(Orders(), "formulas.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sheet = FirstSheet(document);

            // The leading = and outer spaces go; no result means no value and no type.
            var amount = CellOf(sheet, "D2");
            Assert.Equal("B2*C2", amount.CellFormula!.Text);
            Assert.Null(amount.CellValue);
            Assert.Null(amount.DataType);
            Assert.Equal("MAX(D2:D5)", CellOf(sheet, "D8").CellFormula!.Text);

            var total = CellOf(sheet, "D7");
            Assert.Equal("SUM(Amounts)", total.CellFormula!.Text);
            Assert.Equal("70.3", total.CellValue!.Text);
            Assert.Null(total.DataType);

            // A text result stays inline (t="str"), not in the shared strings.
            var status = CellOf(sheet, "D9");
            Assert.Equal(CellValues.String, status.DataType!.Value);
            Assert.Equal("Over budget", status.CellValue!.Text);
            Assert.DoesNotContain(document.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>(),
                s => s.InnerText == "Over budget");

            // A formula cell keeps its style.
            Assert.NotNull(CellOf(sheet, "D7").StyleIndex);

            var calc = document.WorkbookPart.Workbook.GetFirstChild<CalculationProperties>()!;
            Assert.True(calc.FullCalculationOnLoad!.Value);
        }
    }

    [Fact]
    public void WorkbooksWithoutFormulasAskForNoCalculation()
    {
        var book = new XlsxWorkbook();
        book.AddSheet("Plain").Set(1, 1, 1.0);
        var path = Save(book, "no-formulas-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            Assert.Null(document.WorkbookPart!.Workbook.GetFirstChild<CalculationProperties>());
        }
    }

    [Fact]
    public void AFormulaReplacesAValueAndAValueReplacesTheFormulaResult()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.Set(1, 1, "old text");
        sheet.SetFormula(1, 1, "1+1");
        sheet.SetFormula(2, 1, "2+2", 4.0);
        var path = Save(book, "formula-replace-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var a1 = CellOf(FirstSheet(document), "A1");
            Assert.Equal("1+1", a1.CellFormula!.Text);
            Assert.Null(a1.CellValue);
            Assert.Equal("4", CellOf(FirstSheet(document), "A2").CellValue!.Text);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("=")]
    [InlineData("   ")]
    [InlineData(" = ")]
    public void EmptyFormulasAreRejected(string formula)
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        Assert.Throws<ArgumentException>(() => sheet.SetFormula(1, 1, formula));
    }

    [Fact]
    public void FormulasFollowExcelsLengthLimit()
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        sheet.SetFormula(1, 1, new string('1', 8192));
        Assert.Throws<ArgumentException>(() => sheet.SetFormula(1, 1, new string('1', 8193)));
    }

    [Fact]
    public void FormulasDoNotDependOnTheThreadCulture()
    {
        var old = System.Threading.Thread.CurrentThread.CurrentCulture;
        try
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");
            var book = new XlsxWorkbook();
            var sheet = book.AddSheet("One");
            sheet.SetFormula(1, 1, "1.5*2", 3.0);
            sheet.Highlight(2, 1, 3, 1, XlsxRule.Compare(XlsxComparison.GreaterThan, 2.5), Green);
            sheet.Validate(2, 2, 3, 2, XlsxValidation.Decimal(XlsxComparison.Between, 0.5, 1.25));
            var path = Save(book, "culture-rules-check.xlsx");
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                var sheetXml = FirstSheet(document);
                Assert.Equal("2.5", sheetXml.Descendants<ConditionalFormattingRule>().Single().GetFirstChild<Formula>()!.Text);
                var validation = sheetXml.Descendants<DataValidation>().Single();
                Assert.Equal("0.5", validation.Formula1!.Text);
                Assert.Equal("1.25", validation.Formula2!.Text);
            }
        }
        finally
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = old;
        }
    }

    #endregion

    #region Data validation

    /// <summary>
    /// A request form: every kind of validation once, with hints and error messages.
    /// </summary>
    private static XlsxWorkbook RequestForm()
    {
        var book = new XlsxWorkbook { Title = "Data validation" };
        var sheet = book.AddSheet("Request");
        var lists = book.AddSheet("Lists");
        var departments = new[] { "Sales", "Finance", "Research", "Support", "Production" };
        for (var i = 0; i < departments.Length; i++)
        {
            lists.Set(i + 1, 1, departments[i]);
        }

        book.DefineName("Departments", lists, 1, 1, departments.Length, 1);
        lists.Visibility = XlsxSheetVisibility.Hidden;

        var labels = new[] { "Priority", "Department", "Quantity", "Discount", "Delivery date", "Code", "E-mail", "Rating", "Note" };
        for (var i = 0; i < labels.Length; i++)
        {
            sheet.Set(i + 1, 1, labels[i]);
        }

        sheet.ColumnWidth(1, 16);
        sheet.ColumnWidth(2, 24);
        sheet.Format(1, 1, labels.Length, 1, (ref XlsxStyle s) => s.Bold = true);
        sheet.Frame(1, 2, labels.Length, 2, new XlsxBorder("thin", "#808080"), new XlsxBorder("thin", "#D0D0D0"));

        var priority = XlsxValidation.List("Low", "Normal", "High");
        priority.InputTitle = "Priority";
        priority.InputMessage = "Pick a priority from the list.";
        sheet.Validate(1, 2, 1, 2, priority);
        sheet.Validate(2, 2, 2, 2, XlsxValidation.ListFrom("=Departments"));

        var quantity = XlsxValidation.WholeNumber(XlsxComparison.Between, 1, 100);
        quantity.ErrorTitle = "Quantity";
        quantity.ErrorMessage = "A whole number from 1 to 100.";
        sheet.Validate(3, 2, 3, 2, quantity);

        var discount = XlsxValidation.Decimal(XlsxComparison.LessThanOrEqual, 0.25);
        discount.ErrorStyle = XlsxErrorStyle.Warning;
        discount.ErrorMessage = "Discounts above 25% need approval.";
        sheet.Validate(4, 2, 4, 2, discount);
        sheet.Format(4, 2, (ref XlsxStyle s) => s.NumberFormat = "0%");

        sheet.Validate(5, 2, 5, 2, XlsxValidation.Date(XlsxComparison.Between, new DateTime(2026, 1, 1), new DateTime(2026, 12, 31)));
        sheet.Validate(6, 2, 6, 2, XlsxValidation.TextLength(XlsxComparison.Equal, 6));

        var mail = XlsxValidation.Custom("ISNUMBER(SEARCH(\"@\",B7))");
        mail.ErrorStyle = XlsxErrorStyle.Information;
        mail.AllowBlank = false;
        sheet.Validate(7, 2, 7, 2, mail);

        var rating = XlsxValidation.List("1", "2", "3", "4", "5");
        rating.ShowDropDown = false;
        sheet.Validate(8, 2, 8, 2, rating);
        sheet.Set(9, 2, "Type 1–5 into Rating: it has no arrow.");
        return book;
    }

    [Fact]
    public void EveryKindOfValidation()
    {
        var path = Save(RequestForm(), "validation.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sheet = FirstSheet(document);
            var list = sheet.GetFirstChild<DataValidations>()!;
            Assert.Equal(8u, list.Count!.Value);
            var rules = list.Elements<DataValidation>().ToDictionary(v => v.SequenceOfReferences!.InnerText!);

            Assert.Equal(DataValidationValues.List, rules["B1"].Type!.Value);
            Assert.Equal("\"Low,Normal,High\"", rules["B1"].Formula1!.Text);
            Assert.Equal("Pick a priority from the list.", rules["B1"].Prompt!.Value);
            Assert.True(rules["B1"].ShowInputMessage!.Value);
            Assert.True(rules["B1"].ShowErrorMessage!.Value);
            Assert.Null(rules["B1"].Operator);

            Assert.Equal("Departments", rules["B2"].Formula1!.Text);

            Assert.Equal(DataValidationValues.Whole, rules["B3"].Type!.Value);
            Assert.Equal(DataValidationOperatorValues.Between, rules["B3"].Operator!.Value);
            Assert.Equal("1", rules["B3"].Formula1!.Text);
            Assert.Equal("100", rules["B3"].Formula2!.Text);
            Assert.Equal("A whole number from 1 to 100.", rules["B3"].Error!.Value);

            Assert.Equal(DataValidationValues.Decimal, rules["B4"].Type!.Value);
            Assert.Equal(DataValidationOperatorValues.LessThanOrEqual, rules["B4"].Operator!.Value);
            Assert.Equal(DataValidationErrorStyleValues.Warning, rules["B4"].ErrorStyle!.Value);
            Assert.Null(rules["B4"].Formula2);

            Assert.Equal(DataValidationValues.Date, rules["B5"].Type!.Value);
            Assert.Equal(new DateTime(2026, 1, 1).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), rules["B5"].Formula1!.Text);
            Assert.Equal(new DateTime(2026, 12, 31).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), rules["B5"].Formula2!.Text);

            Assert.Equal(DataValidationValues.TextLength, rules["B6"].Type!.Value);
            Assert.Equal(DataValidationOperatorValues.Equal, rules["B6"].Operator!.Value);

            Assert.Equal(DataValidationValues.Custom, rules["B7"].Type!.Value);
            Assert.Equal("ISNUMBER(SEARCH(\"@\",B7))", rules["B7"].Formula1!.Text);
            Assert.Equal(DataValidationErrorStyleValues.Information, rules["B7"].ErrorStyle!.Value);
            Assert.Null(rules["B7"].AllowBlank);

            // showDropDown="1" hides the arrow; only the rating list asks for that.
            Assert.True(rules["B8"].ShowDropDown!.Value);
            Assert.All(rules.Where(r => r.Key != "B8"), r => Assert.Null(r.Value.ShowDropDown));
        }
    }

    [Fact]
    public void ListValuesFollowExcelsLimits()
    {
        Assert.Throws<ArgumentException>(() => XlsxValidation.List("a,b"));
        XlsxValidation.List(Enumerable.Repeat("abcd", 51).ToArray());
        Assert.Throws<ArgumentException>(() => XlsxValidation.List(Enumerable.Repeat("abcde", 43).ToArray()));
    }

    [Fact]
    public void RangeComparisonsNeedTwoValues()
    {
        Assert.Throws<ArgumentException>(() => XlsxValidation.WholeNumber(XlsxComparison.Between, 1));
        Assert.Throws<ArgumentException>(() => XlsxValidation.Date(XlsxComparison.NotBetween, DateTime.Today));
        Assert.Throws<ArgumentException>(() => XlsxRule.Compare(XlsxComparison.Between, 1));
        Assert.Throws<ArgumentException>(() => XlsxRule.CompareFormula(XlsxComparison.NotBetween, "A1"));

        // A second value on a one-sided comparison is ignored.
        XlsxValidation.WholeNumber(XlsxComparison.GreaterThan, 1, 5);
        XlsxRule.Compare(XlsxComparison.LessThan, 1, 5);
    }

    [Fact]
    public void TooLongValidationTextsStopTheSaveAndLeaveNoFile()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        var validation = XlsxValidation.WholeNumber(XlsxComparison.GreaterThan, 0);
        validation.InputTitle = new string('x', 33);
        sheet.Validate(1, 1, 1, 1, validation);
        var folder = XlsxTestKit.NewFolder();
        var path = Path.Combine(folder, "long.xlsx");

        Assert.Throws<InvalidOperationException>(() => book.Save(path));
        Assert.Empty(Directory.GetFiles(folder));

        validation.InputTitle = new string('x', 32);
        validation.InputMessage = new string('y', 255);
        validation.ErrorMessage = new string('z', 255);
        book.Save(path);
        Assert.Empty(XlsxTestKit.Validate(path));
    }

    #endregion

    #region Conditional formatting

    /// <summary>
    /// Exam results with every kind of rule, a few data bars and icon sets.
    /// </summary>
    private static XlsxWorkbook Results()
    {
        var book = new XlsxWorkbook { Title = "Conditional formatting" };
        var sheet = book.AddSheet("Results");
        var headers = new[] { "Student", "Exam 1", "Exam 2", "Exam 3", "Average", "Status", "Bar", "Trend", "Rating", "Note" };
        for (var c = 0; c < headers.Length; c++)
        {
            sheet.Set(1, c + 1, headers[c]);
        }

        sheet.Format(1, 1, 1, headers.Length, (ref XlsxStyle s) => { s.Bold = true; s.Fill = "#DCE6F1"; });
        var names = new[] { "Ada", "Bora", "Cem", "Deniz", "Ece", "Fırat", "Gül", "Hakan", "Ada", "Irmak" };
        var random = new Random(7);
        for (var r = 0; r < names.Length; r++)
        {
            var row = r + 2;
            sheet.Set(row, 1, names[r]);
            var sum = 0.0;
            for (var c = 2; c <= 4; c++)
            {
                var score = 35 + random.Next(65);
                sum += score;
                sheet.Set(row, c, score);
            }

            sheet.SetFormula(row, 5, "AVERAGE(B" + row + ":D" + row + ")", Math.Round(sum / 3, 2));
            sheet.SetFormula(row, 6, "IF(E" + row + ">=50,\"Passed\",\"Late\")", sum / 3 >= 50 ? "Passed" : "Late");
            sheet.SetFormula(row, 7, "E" + row, sum / 3);
            sheet.SetFormula(row, 8, "D" + row + "-B" + row, 0.0);
            sheet.SetFormula(row, 9, "E" + row, sum / 3);
            if (r % 3 == 0)
            {
                sheet.Set(row, 10, r == 0 ? "say \"hi\"" : "Retake " + r);
            }
        }

        var last = names.Length + 1;
        sheet.Format(2, 5, last, 9, (ref XlsxStyle s) => s.NumberFormat = "0.0");
        sheet.ColumnWidth(10, 14);

        sheet.Highlight(2, 2, last, 4, XlsxRule.Compare(XlsxComparison.GreaterThanOrEqual, 85), Green);
        sheet.Highlight(2, 2, last, 4, XlsxRule.Compare(XlsxComparison.Between, 50, 60), Yellow);
        sheet.Highlight(2, 2, last, 4, XlsxRule.CompareFormula(XlsxComparison.LessThan, "$E$2"), new XlsxStyle { Italic = true });
        sheet.Highlight(2, 6, last, 6, XlsxRule.TextContains("late"), Red);
        sheet.Highlight(2, 10, last, 10, XlsxRule.TextBeginsWith("Retake"), new XlsxStyle { Bold = true });
        sheet.Highlight(2, 10, last, 10, XlsxRule.TextEndsWith("\"hi\""), new XlsxStyle { Strikethrough = true });
        sheet.Highlight(2, 10, last, 10, XlsxRule.TextNotContains("e"), new XlsxStyle { FontColor = "#7030A0" });
        sheet.Highlight(2, 10, last, 10, XlsxRule.Blanks(), new XlsxStyle { Fill = "#F2F2F2" });
        sheet.Highlight(2, 5, last, 5, XlsxRule.Top(3), new XlsxStyle { Bold = true, Bottom = new XlsxBorder("medium", "#006100") });
        sheet.Highlight(2, 5, last, 5, XlsxRule.BottomValues(10, percent: true), Red);
        sheet.Highlight(2, 2, last, 2, XlsxRule.AboveAverage(), new XlsxStyle { Underline = true });
        sheet.Highlight(2, 3, last, 3, XlsxRule.BelowAverage(), new XlsxStyle { NumberFormat = "0.0\" low\"" });
        sheet.Highlight(2, 1, last, 1, XlsxRule.Duplicates(), Yellow);
        sheet.Highlight(2, 4, last, 4, XlsxRule.Unique(), new XlsxStyle { DoubleUnderline = true });
        sheet.Highlight(2, 1, last, 1, XlsxRule.Formula("$F2=\"Late\""), new XlsxStyle { FontColor = "#9C0006" });
        sheet.Highlight(2, 9, last, 9, XlsxRule.NotBlanks(), new XlsxStyle { Fill = "#FFFFFF" });
        sheet.DataBar(2, 7, last, 7, "#638EC6");
        sheet.IconSet(2, 8, last, 8, XlsxIconSet.ThreeArrows);
        sheet.IconSet(2, 9, last, 9, XlsxIconSet.FiveRating, showValue: false);
        sheet.IconSet(2, 5, last, 5, XlsxIconSet.FourTrafficLights, reverse: true);
        sheet.ColorScale(2, 2, last, 2, "#F8696B", null, "#63BE7B");

        // The three kinds of trailing parts together, to check their order.
        sheet.Validate(2, 2, last, 4, XlsxValidation.WholeNumber(XlsxComparison.Between, 0, 100));
        sheet.LinkUrl(last + 2, 1, "https://example.com/results", "Results online");
        return book;
    }

    [Fact]
    public void EveryKindOfRule()
    {
        var path = Save(Results(), "conditional.xlsx", keep: true);
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var sheet = FirstSheet(document);
            var blocks = sheet.Elements<ConditionalFormatting>().ToList();
            var rules = blocks.Select(b => b.GetFirstChild<ConditionalFormattingRule>()!).ToList();
            Assert.Equal(21, rules.Count);

            // Priority follows the order of adding: 1, 2, 3 …
            Assert.Equal(Enumerable.Range(1, 21), rules.Select(r => r.Priority!.Value));
            Assert.Equal("B2:D11", blocks[0].SequenceOfReferences!.InnerText);

            Assert.Equal(ConditionalFormatValues.CellIs, rules[0].Type!.Value);
            Assert.Equal(ConditionalFormattingOperatorValues.GreaterThanOrEqual, rules[0].Operator!.Value);
            Assert.Equal("85", rules[0].GetFirstChild<Formula>()!.Text);
            Assert.Equal(new[] { "50", "60" }, rules[1].Elements<Formula>().Select(f => f.Text).ToArray());
            Assert.Equal("$E$2", rules[2].GetFirstChild<Formula>()!.Text);

            // Text rules: the attribute keeps the text, the formula searches from the top-left cell, quotes doubled.
            Assert.Equal(ConditionalFormatValues.ContainsText, rules[3].Type!.Value);
            Assert.Equal("late", rules[3].Text!.Value);
            Assert.Equal("NOT(ISERROR(SEARCH(\"late\",F2)))", rules[3].GetFirstChild<Formula>()!.Text);
            Assert.Equal("LEFT(J2,LEN(\"Retake\"))=\"Retake\"", rules[4].GetFirstChild<Formula>()!.Text);
            Assert.Equal("RIGHT(J2,LEN(\"\"\"hi\"\"\"))=\"\"\"hi\"\"\"", rules[5].GetFirstChild<Formula>()!.Text);
            Assert.Equal(ConditionalFormatValues.NotContainsText, rules[6].Type!.Value);
            Assert.Equal(ConditionalFormattingOperatorValues.NotContains, rules[6].Operator!.Value);
            Assert.Equal("LEN(TRIM(J2))=0", rules[7].GetFirstChild<Formula>()!.Text);

            Assert.Equal(ConditionalFormatValues.Top10, rules[8].Type!.Value);
            Assert.Equal(3u, rules[8].Rank!.Value);
            Assert.Null(rules[8].Bottom);
            Assert.True(rules[9].Bottom!.Value);
            Assert.True(rules[9].Percent!.Value);
            Assert.Equal(10u, rules[9].Rank!.Value);

            Assert.Equal(ConditionalFormatValues.AboveAverage, rules[10].Type!.Value);
            Assert.Null(rules[10].AboveAverage);
            Assert.False(rules[11].AboveAverage!.Value);
            Assert.Equal(ConditionalFormatValues.DuplicateValues, rules[12].Type!.Value);
            Assert.Equal(ConditionalFormatValues.UniqueValues, rules[13].Type!.Value);
            Assert.Equal(ConditionalFormatValues.Expression, rules[14].Type!.Value);
            Assert.Equal("$F2=\"Late\"", rules[14].GetFirstChild<Formula>()!.Text);
            Assert.Equal("LEN(TRIM(I2))>0", rules[15].GetFirstChild<Formula>()!.Text);

            Assert.Equal(ConditionalFormatValues.DataBar, rules[16].Type!.Value);
            Assert.Equal("FF638EC6", rules[16].GetFirstChild<DataBar>()!.GetFirstChild<Color>()!.Rgb!.Value);

            var arrows = rules[17].GetFirstChild<IconSet>()!;
            Assert.Equal(IconSetValues.ThreeArrows, arrows.IconSetValue!.Value);
            Assert.Equal(new[] { "0", "33", "67" }, arrows.Elements<ConditionalFormatValueObject>().Select(v => v.Val!.Value).ToArray());
            var rating = rules[18].GetFirstChild<IconSet>()!;
            Assert.Equal(IconSetValues.FiveRating, rating.IconSetValue!.Value);
            Assert.False(rating.ShowValue!.Value);
            Assert.Equal(new[] { "0", "20", "40", "60", "80" }, rating.Elements<ConditionalFormatValueObject>().Select(v => v.Val!.Value).ToArray());
            var lights = rules[19].GetFirstChild<IconSet>()!;
            Assert.True(lights.Reverse!.Value);
            Assert.Equal(new[] { "0", "25", "50", "75" }, lights.Elements<ConditionalFormatValueObject>().Select(v => v.Val!.Value).ToArray());
            Assert.Equal(ConditionalFormatValues.ColorScale, rules[20].Type!.Value);

            // Rules without a look have no dxfId; rules with the same look share one.
            Assert.All(rules.Skip(16), r => Assert.Null(r.FormatId));
            Assert.Equal(rules[3].FormatId!.Value, rules[9].FormatId!.Value);
            Assert.Equal(rules[1].FormatId!.Value, rules[12].FormatId!.Value);
            var dxfs = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet.DifferentialFormats!;
            var used = rules.Where(r => r.FormatId != null).Select(r => r.FormatId!.Value).Distinct().Count();
            Assert.Equal((uint)used, dxfs.Count!.Value);
            Assert.All(rules.Where(r => r.FormatId != null), r => Assert.InRange(r.FormatId!.Value, 0u, dxfs.Count.Value - 1));

            // A differential fill puts its colour in bgColor; the number format is declared with the cell formats.
            var red = dxfs.Elements<DifferentialFormat>().ElementAt((int)rules[3].FormatId!.Value);
            Assert.Equal("FFFFC7CE", red.Fill!.PatternFill!.BackgroundColor!.Rgb!.Value);
            Assert.Equal("FF9C0006", red.Font!.Color!.Rgb!.Value);
            var low = dxfs.Elements<DifferentialFormat>().ElementAt((int)rules[11].FormatId!.Value);
            var formats = document.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats!.Elements<NumberingFormat>();
            Assert.Contains(formats, f => f.NumberFormatId!.Value == low.NumberingFormat!.NumberFormatId!.Value && f.FormatCode == "0.0\" low\"");
            var top = dxfs.Elements<DifferentialFormat>().ElementAt((int)rules[8].FormatId!.Value);
            Assert.Equal(BorderStyleValues.Medium, top.Border!.BottomBorder!.Style!.Value);
            Assert.Null(top.Border.LeftBorder);
        }
    }

    [Fact]
    public void ConditionalFormatsAreReproducible()
    {
        var folder = XlsxTestKit.NewFolder();
        var a = Path.Combine(folder, "a.xlsx");
        var b = Path.Combine(folder, "b.xlsx");
        var first = Results();
        first.Created = XlsxTestKit.FixedTime;
        first.Save(a);
        var second = Results();
        second.Created = XlsxTestKit.FixedTime;
        second.Save(b);
        Assert.Equal(File.ReadAllBytes(a), File.ReadAllBytes(b));

        // Saving the same workbook twice gives the same bytes too: no state is left over from the first save.
        first.Save(b);
        Assert.Equal(File.ReadAllBytes(a), File.ReadAllBytes(b));
    }

    [Fact]
    public void RangesAreChecked()
    {
        var sheet = new XlsxWorkbook().AddSheet("One");
        Assert.Throws<ArgumentException>(() => sheet.Highlight(5, 1, 1, 1, XlsxRule.Duplicates(), Green));
        Assert.Throws<ArgumentException>(() => sheet.DataBar(1, 3, 5, 1, "#638EC6"));
        Assert.Throws<ArgumentException>(() => sheet.Validate(2, 1, 1, 1, XlsxValidation.List("a")));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.IconSet(0, 1, 5, 1, XlsxIconSet.ThreeFlags));
        Assert.Throws<ArgumentException>(() => sheet.DataBar(1, 1, 5, 1, "blue"));
    }

    [Fact]
    public void RankLimitsAreKept()
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        sheet.Highlight(1, 1, 5, 1, XlsxRule.Top(5000), Green);
        sheet.Highlight(1, 1, 5, 1, XlsxRule.Top(250, percent: true), Green);
        sheet.Highlight(1, 1, 5, 1, XlsxRule.BottomValues(0), Green);
        var path = Save(book, "rank-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var ranks = FirstSheet(document).Descendants<ConditionalFormattingRule>().Select(r => r.Rank!.Value).ToArray();
            Assert.Equal(new uint[] { 1000, 100, 1 }, ranks);
        }
    }

    [Theory]
    [InlineData(XlsxIconSet.ThreeTrafficLights, "3TrafficLights1", 3)]
    [InlineData(XlsxIconSet.ThreeTrafficLightsFramed, "3TrafficLights2", 3)]
    [InlineData(XlsxIconSet.ThreeSigns, "3Signs", 3)]
    [InlineData(XlsxIconSet.ThreeArrowsGray, "3ArrowsGray", 3)]
    [InlineData(XlsxIconSet.ThreeFlags, "3Flags", 3)]
    [InlineData(XlsxIconSet.ThreeSymbols, "3Symbols", 3)]
    [InlineData(XlsxIconSet.ThreeSymbolsPlain, "3Symbols2", 3)]
    [InlineData(XlsxIconSet.FourArrows, "4Arrows", 4)]
    [InlineData(XlsxIconSet.FourArrowsGray, "4ArrowsGray", 4)]
    [InlineData(XlsxIconSet.FourRedToBlack, "4RedToBlack", 4)]
    [InlineData(XlsxIconSet.FourRating, "4Rating", 4)]
    [InlineData(XlsxIconSet.FiveArrows, "5Arrows", 5)]
    [InlineData(XlsxIconSet.FiveArrowsGray, "5ArrowsGray", 5)]
    [InlineData(XlsxIconSet.FiveQuarters, "5Quarters", 5)]
    public void EveryIconSetIsValid(XlsxIconSet icons, string name, int count)
    {
        var book = new XlsxWorkbook();
        var sheet = book.AddSheet("One");
        for (var r = 1; r <= 10; r++)
        {
            sheet.Set(r, 1, r);
        }

        sheet.IconSet(1, 1, 10, 1, icons);
        var path = Save(book, "icons-check.xlsx");
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var set = FirstSheet(document).Descendants<IconSet>().Single();
            Assert.Equal(name, set.IconSetValue!.InnerText);
            Assert.Equal(count, set.Elements<ConditionalFormatValueObject>().Count());
        }
    }

    #endregion
}
