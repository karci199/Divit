#nullable enable
namespace Divit.Docx;

using System;
using System.Collections.Generic;

/// <summary>
/// A piece of an equation (Office Math), built from the factory methods and nested:
/// <c>DocxMath.Fraction("2", DocxMath.Row(DocxMath.Sup("D", "2")))</c>. Text is set in italics as a variable;
/// <see cref="Plain"/> sets units and words upright. A string converts to <see cref="Text"/>.
/// </summary>
/// <remarks>
/// Word lays out and edits the equation like one typed in its equation editor, in the Cambria Math font.
/// </remarks>
public sealed class DocxMath
{
    private DocxMath(DocxMathKind kind, string? text = null, params DocxMath[] parts)
    {
        Kind = kind;
        Content = text;
        Parts = parts;
    }

    internal DocxMathKind Kind { get; }

    /// <summary>
    /// The text of a run, the name of a function, or the characters of a bracket or operator.
    /// </summary>
    internal string? Content { get; }

    /// <summary>
    /// The second character: the closing bracket.
    /// </summary>
    internal string? Close { get; private set; }

    internal IReadOnlyList<DocxMath> Parts { get; }

    /// <summary>
    /// Text set as variables, in italics: <c>x</c>, <c>λ</c>; digits and operators stay upright.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Text(string text) => new DocxMath(DocxMathKind.Text, text ?? string.Empty);

    /// <summary>
    /// Text set upright: units, words, abbreviations (<c>dB</c>, <c>max</c>).
    /// </summary>
    /// <param name="text">The text.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Plain(string text) => new DocxMath(DocxMathKind.Plain, text ?? string.Empty);

    /// <summary>
    /// Pieces one after another.
    /// </summary>
    /// <param name="parts">The pieces.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Row(params DocxMath[] parts) => new DocxMath(DocxMathKind.Row, null, Check(parts));

    /// <summary>
    /// A fraction: numerator over denominator.
    /// </summary>
    /// <param name="numerator">Above the line.</param>
    /// <param name="denominator">Below the line.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Fraction(DocxMath numerator, DocxMath denominator) =>
        new DocxMath(DocxMathKind.Fraction, null, Check(numerator, denominator));

    /// <summary>
    /// A raised index: a power, <c>D²</c>.
    /// </summary>
    /// <param name="base">The base.</param>
    /// <param name="superscript">The raised part.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Sup(DocxMath @base, DocxMath superscript) =>
        new DocxMath(DocxMathKind.Sup, null, Check(@base, superscript));

    /// <summary>
    /// A lowered index: <c>Gₜ</c>.
    /// </summary>
    /// <param name="base">The base.</param>
    /// <param name="subscript">The lowered part.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Sub(DocxMath @base, DocxMath subscript) =>
        new DocxMath(DocxMathKind.Sub, null, Check(@base, subscript));

    /// <summary>
    /// A lowered and a raised index on one base.
    /// </summary>
    /// <param name="base">The base.</param>
    /// <param name="subscript">The lowered part.</param>
    /// <param name="superscript">The raised part.</param>
    /// <returns>The piece.</returns>
    public static DocxMath SubSup(DocxMath @base, DocxMath subscript, DocxMath superscript) =>
        new DocxMath(DocxMathKind.SubSup, null, Check(@base, subscript, superscript));

    /// <summary>
    /// A square root, or with a degree any root.
    /// </summary>
    /// <param name="radicand">Under the root.</param>
    /// <param name="degree">The degree (3 for a cube root); none for a square root.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Root(DocxMath radicand, DocxMath? degree = null) =>
        degree == null ? new DocxMath(DocxMathKind.Root, null, Check(radicand)) : new DocxMath(DocxMathKind.Root, null, Check(radicand, degree));

    /// <summary>
    /// Brackets that grow with their content: (), [], {}, |  |.
    /// </summary>
    /// <param name="content">Inside the brackets.</param>
    /// <param name="open">The opening bracket.</param>
    /// <param name="close">The closing bracket.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Brackets(DocxMath content, string open = "(", string close = ")") =>
        new DocxMath(DocxMathKind.Brackets, open ?? string.Empty, Check(content)) { Close = close ?? string.Empty };

    /// <summary>
    /// A sum with its limits below and above: ∑.
    /// </summary>
    /// <param name="body">What is summed.</param>
    /// <param name="lower">The lower limit; none when empty.</param>
    /// <param name="upper">The upper limit; none when empty.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Sum(DocxMath body, DocxMath? lower = null, DocxMath? upper = null) =>
        Operator("∑", body, lower, upper);

    /// <summary>
    /// An integral with its limits at the side: ∫.
    /// </summary>
    /// <param name="body">What is integrated, with its <c>dx</c>.</param>
    /// <param name="lower">The lower limit; none when empty.</param>
    /// <param name="upper">The upper limit; none when empty.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Integral(DocxMath body, DocxMath? lower = null, DocxMath? upper = null) =>
        Operator("∫", body, lower, upper);

    /// <summary>
    /// A large operator with limits, such as ∏ or ∮.
    /// </summary>
    /// <param name="symbol">The operator's character.</param>
    /// <param name="body">Its operand.</param>
    /// <param name="lower">The lower limit; none when empty.</param>
    /// <param name="upper">The upper limit; none when empty.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Operator(string symbol, DocxMath body, DocxMath? lower = null, DocxMath? upper = null) =>
        new DocxMath(DocxMathKind.Operator, symbol, Check(body, lower ?? Row(), upper ?? Row())) { HasLower = lower != null, HasUpper = upper != null };

    /// <summary>
    /// A function with its argument: <c>sin x</c>, <c>log₁₀ x</c>; the name is set upright.
    /// </summary>
    /// <param name="name">The name, for example <c>sin</c>, or a piece such as <c>Sub(Plain("log"), "10")</c>.</param>
    /// <param name="argument">The argument.</param>
    /// <returns>The piece.</returns>
    public static DocxMath Function(DocxMath name, DocxMath argument) =>
        new DocxMath(DocxMathKind.Function, null, Check(name.Kind == DocxMathKind.Text ? Plain(name.Content!) : name, argument));

    /// <summary>
    /// Text as a variable, as <see cref="Text"/> makes it.
    /// </summary>
    /// <param name="text">The text.</param>
    public static implicit operator DocxMath(string text) => Text(text);

    /// <summary>
    /// Whether an operator shows its lower and upper limits.
    /// </summary>
    internal bool HasLower { get; private set; }

    internal bool HasUpper { get; private set; }

    /// <summary>
    /// The plain text of the equation, as a fallback and for tests: brackets and operators in line.
    /// </summary>
    /// <returns>The text.</returns>
    public override string ToString()
    {
        switch (Kind)
        {
            case DocxMathKind.Text:
            case DocxMathKind.Plain:
                return Content!;
            case DocxMathKind.Fraction:
                return "(" + Parts[0] + ")/(" + Parts[1] + ")";
            case DocxMathKind.Sup:
                return Parts[0] + "^" + Parts[1];
            case DocxMathKind.Sub:
                return Parts[0] + "_" + Parts[1];
            case DocxMathKind.SubSup:
                return Parts[0] + "_" + Parts[1] + "^" + Parts[2];
            case DocxMathKind.Root:
                return Parts.Count > 1 ? "root(" + Parts[1] + ")(" + Parts[0] + ")" : "√(" + Parts[0] + ")";
            case DocxMathKind.Brackets:
                return Content + Parts[0] + Close;
            case DocxMathKind.Operator:
                return Content + (HasLower ? "_" + Parts[1] : string.Empty) + (HasUpper ? "^" + Parts[2] : string.Empty) + " " + Parts[0];
            case DocxMathKind.Function:
                return Parts[0] + " " + Parts[1];
            default:
                return string.Concat(Parts);
        }
    }

    private static DocxMath[] Check(params DocxMath[] parts)
    {
        foreach (var part in parts)
        {
            if (part == null)
            {
                throw new ArgumentNullException(nameof(parts));
            }
        }

        return parts;
    }
}

/// <summary>
/// The kinds of equation pieces, one per Office Math element.
/// </summary>
internal enum DocxMathKind
{
    Text,
    Plain,
    Row,
    Fraction,
    Sup,
    Sub,
    SubSup,
    Root,
    Brackets,
    Operator,
    Function,
}
