#nullable enable
namespace Divit.Xlsx;

using Divit.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;

/// <summary>
/// How a cell value is compared, in conditional formatting rules and data validation.
/// </summary>
public enum XlsxComparison
{
    /// <summary>
    /// Between the two values, both included.
    /// </summary>
    Between,

    /// <summary>
    /// Outside the two values.
    /// </summary>
    NotBetween,

    /// <summary>
    /// Equal to the value.
    /// </summary>
    Equal,

    /// <summary>
    /// Not equal to the value.
    /// </summary>
    NotEqual,

    /// <summary>
    /// Greater than the value.
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Less than the value.
    /// </summary>
    LessThan,

    /// <summary>
    /// Greater than or equal to the value.
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// Less than or equal to the value.
    /// </summary>
    LessThanOrEqual,
}

/// <summary>
/// The icons of an icon set; the name tells how many icons it has, from the lowest values to the highest.
/// </summary>
public enum XlsxIconSet
{
    /// <summary>
    /// Red, yellow, green circles.
    /// </summary>
    ThreeTrafficLights,

    /// <summary>
    /// Red, yellow, green lights in a black frame.
    /// </summary>
    ThreeTrafficLightsFramed,

    /// <summary>
    /// Red diamond, yellow triangle, green circle.
    /// </summary>
    ThreeSigns,

    /// <summary>
    /// Red down, yellow sideways, green up arrows.
    /// </summary>
    ThreeArrows,

    /// <summary>
    /// Grey down, sideways and up arrows.
    /// </summary>
    ThreeArrowsGray,

    /// <summary>
    /// Red, yellow, green flags.
    /// </summary>
    ThreeFlags,

    /// <summary>
    /// Cross, exclamation mark, tick in circles.
    /// </summary>
    ThreeSymbols,

    /// <summary>
    /// Cross, exclamation mark, tick without circles.
    /// </summary>
    ThreeSymbolsPlain,

    /// <summary>
    /// Four arrows from down to up.
    /// </summary>
    FourArrows,

    /// <summary>
    /// Four grey arrows from down to up.
    /// </summary>
    FourArrowsGray,

    /// <summary>
    /// Black, grey, pink, red circles.
    /// </summary>
    FourRedToBlack,

    /// <summary>
    /// Signal bars, one to four.
    /// </summary>
    FourRating,

    /// <summary>
    /// Black, red, yellow, green circles.
    /// </summary>
    FourTrafficLights,

    /// <summary>
    /// Five arrows from down to up.
    /// </summary>
    FiveArrows,

    /// <summary>
    /// Five grey arrows from down to up.
    /// </summary>
    FiveArrowsGray,

    /// <summary>
    /// Signal bars, none to four.
    /// </summary>
    FiveRating,

    /// <summary>
    /// Circles from empty to full in quarters.
    /// </summary>
    FiveQuarters,
}

/// <summary>
/// How Excel reacts when a value typed into a validated cell is not allowed.
/// </summary>
public enum XlsxErrorStyle
{
    /// <summary>
    /// The value is refused; the user must retry or cancel.
    /// </summary>
    Stop,

    /// <summary>
    /// The user is warned and may keep the value.
    /// </summary>
    Warning,

    /// <summary>
    /// The user is told and the value is kept.
    /// </summary>
    Information,
}

/// <summary>
/// The condition of a highlight rule (<see cref="XlsxSheet.Highlight"/>): when it holds for a cell, the cell takes
/// the rule's look.
/// </summary>
/// <remarks>
/// Formulas are written as in Excel without the leading <c>=</c>, with English function names and commas, relative
/// to the top-left cell of the range: <c>A2&gt;B2</c> on the range A2:A10 compares every row with column B.
/// </remarks>
public sealed class XlsxRule
{
    private XlsxRule(string type)
    {
        Type = type;
    }

    internal string Type { get; }

    internal string? Operator { get; private set; }

    /// <summary>
    /// Formulas of the rule; for text rules the top-left cell is filled in when the sheet is saved.
    /// </summary>
    internal List<string> Formulas { get; } = new List<string>();

    internal string? Text { get; private set; }

    internal int Rank { get; private set; }

    internal bool Percent { get; private set; }

    internal bool Bottom { get; private set; }

    internal bool Below { get; private set; }

    /// <summary>
    /// Cells whose value compares to one number (or two, for <see cref="XlsxComparison.Between"/> and
    /// <see cref="XlsxComparison.NotBetween"/>).
    /// </summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="value">The value, or the lower end of a range.</param>
    /// <param name="second">The upper end of a range.</param>
    /// <returns>The rule.</returns>
    /// <exception cref="ArgumentException">A range comparison has no second value.</exception>
    public static XlsxRule Compare(XlsxComparison comparison, double value, double? second = null) =>
        CompareFormula(comparison, Number(value), second == null ? null : Number(second.Value));

    /// <summary>
    /// Cells whose value compares to the result of a formula, for example <c>$B$1</c> or <c>AVERAGE($A:$A)</c>.
    /// </summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="formula">The formula, or the lower end of a range.</param>
    /// <param name="second">The upper end of a range.</param>
    /// <returns>The rule.</returns>
    /// <exception cref="ArgumentException">A range comparison has no second formula.</exception>
    public static XlsxRule CompareFormula(XlsxComparison comparison, string formula, string? second = null)
    {
        var rule = new XlsxRule("cellIs") { Operator = OperatorName(comparison) };
        rule.Formulas.Add(XlsxSheet.TrimFormula(formula));
        if (IsRange(comparison))
        {
            if (second == null)
            {
                throw new ArgumentException(Messages.Get(Strings.SecondValueMissing), nameof(second));
            }

            rule.Formulas.Add(XlsxSheet.TrimFormula(second));
        }

        return rule;
    }

    /// <summary>
    /// Cells whose text contains a piece of text, ignoring case.
    /// </summary>
    /// <param name="text">The text to look for.</param>
    /// <returns>The rule.</returns>
    public static XlsxRule TextContains(string text) => TextRule("containsText", "containsText", text);

    /// <summary>
    /// Cells whose text does not contain a piece of text, ignoring case.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <returns>The rule.</returns>
    public static XlsxRule TextNotContains(string text) => TextRule("notContainsText", "notContains", text);

    /// <summary>
    /// Cells whose text begins with a piece of text, ignoring case.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <returns>The rule.</returns>
    public static XlsxRule TextBeginsWith(string text) => TextRule("beginsWith", "beginsWith", text);

    /// <summary>
    /// Cells whose text ends with a piece of text, ignoring case.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <returns>The rule.</returns>
    public static XlsxRule TextEndsWith(string text) => TextRule("endsWith", "endsWith", text);

    /// <summary>
    /// The highest values of the range.
    /// </summary>
    /// <param name="rank">How many values (1–1000), or what percentage of them with <paramref name="percent"/>.</param>
    /// <param name="percent">Whether <paramref name="rank"/> is a percentage (1–100).</param>
    /// <returns>The rule.</returns>
    public static XlsxRule Top(int rank, bool percent = false) =>
        new XlsxRule("top10") { Rank = Math.Max(1, Math.Min(percent ? 100 : 1000, rank)), Percent = percent };

    /// <summary>
    /// The lowest values of the range.
    /// </summary>
    /// <param name="rank">How many values (1–1000), or what percentage of them with <paramref name="percent"/>.</param>
    /// <param name="percent">Whether <paramref name="rank"/> is a percentage (1–100).</param>
    /// <returns>The rule.</returns>
    public static XlsxRule BottomValues(int rank, bool percent = false)
    {
        var rule = Top(rank, percent);
        rule.Bottom = true;
        return rule;
    }

    /// <summary>
    /// Values above the average of the range.
    /// </summary>
    /// <returns>The rule.</returns>
    public static XlsxRule AboveAverage() => new XlsxRule("aboveAverage");

    /// <summary>
    /// Values below the average of the range.
    /// </summary>
    /// <returns>The rule.</returns>
    public static XlsxRule BelowAverage() => new XlsxRule("aboveAverage") { Below = true };

    /// <summary>
    /// Values that occur more than once in the range.
    /// </summary>
    /// <returns>The rule.</returns>
    public static XlsxRule Duplicates() => new XlsxRule("duplicateValues");

    /// <summary>
    /// Values that occur once in the range.
    /// </summary>
    /// <returns>The rule.</returns>
    public static XlsxRule Unique() => new XlsxRule("uniqueValues");

    /// <summary>
    /// Empty cells (or cells holding only spaces).
    /// </summary>
    /// <returns>The rule.</returns>
    public static XlsxRule Blanks() => new XlsxRule("containsBlanks");

    /// <summary>
    /// Cells that are not empty.
    /// </summary>
    /// <returns>The rule.</returns>
    public static XlsxRule NotBlanks() => new XlsxRule("notContainsBlanks");

    /// <summary>
    /// Cells for which a formula is true, for example <c>$C2="Late"</c> to colour whole rows by column C.
    /// </summary>
    /// <param name="formula">The formula.</param>
    /// <returns>The rule.</returns>
    public static XlsxRule Formula(string formula)
    {
        var rule = new XlsxRule("expression");
        rule.Formulas.Add(XlsxSheet.TrimFormula(formula));
        return rule;
    }

    /// <summary>
    /// The formula Excel stores with a text or blank rule; it needs the top-left cell of the range.
    /// </summary>
    internal string? CellFormula(string topLeft)
    {
        var quoted = Text == null ? null : Quote(Text);
        switch (Type)
        {
            case "containsText": return "NOT(ISERROR(SEARCH(" + quoted + "," + topLeft + ")))";
            case "notContainsText": return "ISERROR(SEARCH(" + quoted + "," + topLeft + "))";
            case "beginsWith": return "LEFT(" + topLeft + ",LEN(" + quoted + "))=" + quoted;
            case "endsWith": return "RIGHT(" + topLeft + ",LEN(" + quoted + "))=" + quoted;
            case "containsBlanks": return "LEN(TRIM(" + topLeft + "))=0";
            case "notContainsBlanks": return "LEN(TRIM(" + topLeft + "))>0";
            default: return null;
        }
    }

    internal static string OperatorName(XlsxComparison comparison)
    {
        switch (comparison)
        {
            case XlsxComparison.Between: return "between";
            case XlsxComparison.NotBetween: return "notBetween";
            case XlsxComparison.Equal: return "equal";
            case XlsxComparison.NotEqual: return "notEqual";
            case XlsxComparison.GreaterThan: return "greaterThan";
            case XlsxComparison.LessThan: return "lessThan";
            case XlsxComparison.GreaterThanOrEqual: return "greaterThanOrEqual";
            default: return "lessThanOrEqual";
        }
    }

    internal static bool IsRange(XlsxComparison comparison) =>
        comparison == XlsxComparison.Between || comparison == XlsxComparison.NotBetween;

    internal static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    /// <summary>
    /// A text as a formula string literal: in quotes, inner quotes doubled.
    /// </summary>
    internal static string Quote(string text) => "\"" + text.Replace("\"", "\"\"") + "\"";

    private static XlsxRule TextRule(string type, string op, string text) => new XlsxRule(type) { Operator = op, Text = text };
}

/// <summary>
/// What may be typed into a range (<see cref="XlsxSheet.Validate"/>): a list to pick from, numbers, dates or text
/// lengths within limits, or a formula that must be true. Excel checks typed values only; pasted values pass.
/// </summary>
public sealed class XlsxValidation
{
    private XlsxValidation(string type)
    {
        Type = type;
    }

    /// <summary>
    /// Title of the hint shown while the cell is selected; at most 32 characters. No hint when both title and
    /// <see cref="InputMessage"/> are empty.
    /// </summary>
    public string? InputTitle { get; set; }

    /// <summary>
    /// Text of the hint shown while the cell is selected; at most 255 characters.
    /// </summary>
    public string? InputMessage { get; set; }

    /// <summary>
    /// Title of the message shown for a value that is not allowed; at most 32 characters.
    /// </summary>
    public string? ErrorTitle { get; set; }

    /// <summary>
    /// Text of the message shown for a value that is not allowed; at most 255 characters. Excel's own text when empty.
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Whether a value that is not allowed is refused, or only warned about.
    /// </summary>
    public XlsxErrorStyle ErrorStyle { get; set; }

    /// <summary>
    /// Whether an empty cell is allowed.
    /// </summary>
    public bool AllowBlank { get; set; } = true;

    /// <summary>
    /// Whether a list shows its drop-down arrow; without it the user types a value from the list.
    /// </summary>
    public bool ShowDropDown { get; set; } = true;

    internal string Type { get; }

    internal string? Operator { get; private set; }

    internal string Formula1 { get; private set; } = string.Empty;

    internal string? Formula2 { get; private set; }

    /// <summary>
    /// A drop-down list of fixed values.
    /// </summary>
    /// <param name="items">The values; none may contain a comma, and together they hold at most 255 characters.</param>
    /// <returns>The validation.</returns>
    /// <exception cref="ArgumentException">A value holds a comma or the list is too long.</exception>
    public static XlsxValidation List(params string[] items)
    {
        foreach (var item in items)
        {
            if (item.IndexOf(',') >= 0)
            {
                throw new ArgumentException(Messages.Format(Strings.ListItemComma, item), nameof(items));
            }
        }

        var joined = string.Join(",", items);
        if (joined.Length > 255)
        {
            throw new ArgumentException(Messages.Get(Strings.ListTooLong), nameof(items));
        }

        return new XlsxValidation("list") { Formula1 = XlsxRule.Quote(joined) };
    }

    /// <summary>
    /// A drop-down list of the values in a range or a defined name, for example <c>Lists!$A$1:$A$20</c> or
    /// <c>Colours</c>. Longer lists than <see cref="List"/> allows go this way.
    /// </summary>
    /// <param name="source">The range or name.</param>
    /// <returns>The validation.</returns>
    public static XlsxValidation ListFrom(string source) => new XlsxValidation("list") { Formula1 = XlsxSheet.TrimFormula(source) };

    /// <summary>
    /// Whole numbers that compare to one number, or lie between (outside) two.
    /// </summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="value">The value, or the lower end of a range.</param>
    /// <param name="second">The upper end of a range.</param>
    /// <returns>The validation.</returns>
    public static XlsxValidation WholeNumber(XlsxComparison comparison, double value, double? second = null) =>
        Compare("whole", comparison, XlsxRule.Number(value), second == null ? null : XlsxRule.Number(second.Value));

    /// <summary>
    /// Numbers (with fractions) that compare to one number, or lie between (outside) two.
    /// </summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="value">The value, or the lower end of a range.</param>
    /// <param name="second">The upper end of a range.</param>
    /// <returns>The validation.</returns>
    public static XlsxValidation Decimal(XlsxComparison comparison, double value, double? second = null) =>
        Compare("decimal", comparison, XlsxRule.Number(value), second == null ? null : XlsxRule.Number(second.Value));

    /// <summary>
    /// Dates that compare to one date, or lie between (outside) two.
    /// </summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="value">The date, or the start of a range.</param>
    /// <param name="second">The end of a range.</param>
    /// <returns>The validation.</returns>
    public static XlsxValidation Date(XlsxComparison comparison, DateTime value, DateTime? second = null) =>
        Compare("date", comparison, XlsxRule.Number(XlsxSheet.DateNumber(value)),
            second == null ? null : XlsxRule.Number(XlsxSheet.DateNumber(second.Value)));

    /// <summary>
    /// Text whose length compares to one number, or lies between (outside) two.
    /// </summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="length">The length, or the lower end of a range.</param>
    /// <param name="second">The upper end of a range.</param>
    /// <returns>The validation.</returns>
    public static XlsxValidation TextLength(XlsxComparison comparison, int length, int? second = null) =>
        Compare("textLength", comparison, XlsxRule.Number(length), second == null ? null : XlsxRule.Number(second.Value));

    /// <summary>
    /// Values for which a formula is true, relative to the top-left cell: <c>ISNUMBER(SEARCH("@",A2))</c>.
    /// </summary>
    /// <param name="formula">The formula.</param>
    /// <returns>The validation.</returns>
    public static XlsxValidation Custom(string formula) => new XlsxValidation("custom") { Formula1 = XlsxSheet.TrimFormula(formula) };

    /// <summary>
    /// Throws when a message is longer than Excel stores.
    /// </summary>
    internal void CheckLengths()
    {
        Check(InputTitle, 32);
        Check(ErrorTitle, 32);
        Check(InputMessage, 255);
        Check(ErrorMessage, 255);
    }

    private static void Check(string? text, int limit)
    {
        if (text != null && text.Length > limit)
        {
            throw new InvalidOperationException(Messages.Format(Strings.ValidationTextTooLong, limit, text));
        }
    }

    private static XlsxValidation Compare(string type, XlsxComparison comparison, string first, string? second)
    {
        if (XlsxRule.IsRange(comparison) && second == null)
        {
            throw new ArgumentException(Messages.Get(Strings.SecondValueMissing), nameof(second));
        }

        return new XlsxValidation(type)
        {
            Operator = XlsxRule.OperatorName(comparison),
            Formula1 = first,
            Formula2 = XlsxRule.IsRange(comparison) ? second : null,
        };
    }
}
