#nullable enable
namespace Divit.Xlsx;

using System;
using System.Collections.Generic;

/// <summary>
/// A piece of text with its own font, a part of a rich text cell (<see cref="XlsxSheet.SetRichText"/>).
/// </summary>
public readonly struct XlsxRun : IEquatable<XlsxRun>
{
    /// <summary>
    /// Creates a piece of text.
    /// </summary>
    /// <param name="text">The text.</param>
    /// <param name="look">Its font: name, size, bold, italic, underline, strikethrough, super/subscript and colour
    /// apply, the rest of the style does not. <c>null</c> takes the cell's font.</param>
    public XlsxRun(string text, XlsxStyle? look = null)
    {
        Text = text;
        Look = look;
    }

    /// <summary>
    /// The text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// The font of the piece; <c>null</c> for the cell's font.
    /// </summary>
    public XlsxStyle? Look { get; }

    /// <summary>
    /// A piece of text in the cell's font.
    /// </summary>
    /// <param name="text">The text.</param>
    public static implicit operator XlsxRun(string text) => new XlsxRun(text);

    /// <inheritdoc />
    public bool Equals(XlsxRun other) => Text == other.Text && Nullable.Equals(Look, other.Look);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is XlsxRun other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => ((Text?.GetHashCode() ?? 0) * 397) ^ (Look?.GetHashCode() ?? 0);
}

/// <summary>
/// The pieces of a rich text cell; equal pieces give equal texts, so the shared string table stores each once.
/// </summary>
internal sealed class XlsxRichText : IEquatable<XlsxRichText>
{
    private readonly int _hash;

    public XlsxRichText(IReadOnlyList<XlsxRun> runs)
    {
        Runs = runs;
        var hash = runs.Count;
        foreach (var run in runs)
        {
            hash = unchecked((hash * 31) + run.GetHashCode());
        }

        _hash = hash;
    }

    public IReadOnlyList<XlsxRun> Runs { get; }

    /// <summary>
    /// The plain text, all pieces together.
    /// </summary>
    public string Text => string.Concat(System.Linq.Enumerable.Select(Runs, r => r.Text));

    public bool Equals(XlsxRichText? other)
    {
        if (other == null || other._hash != _hash || other.Runs.Count != Runs.Count)
        {
            return false;
        }

        for (var i = 0; i < Runs.Count; i++)
        {
            if (!Runs[i].Equals(other.Runs[i]))
            {
                return false;
            }
        }

        return true;
    }

    public override bool Equals(object? obj) => Equals(obj as XlsxRichText);

    public override int GetHashCode() => _hash;
}
