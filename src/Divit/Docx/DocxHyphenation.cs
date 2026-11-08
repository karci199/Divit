#nullable enable
namespace Divit.Docx;

using System;
using System.Collections.Generic;

/// <summary>
/// Where words may be broken at the end of a line, by the syllable rules of a language. Turkish syllables are
/// regular, so no dictionary is needed: every syllable has one vowel, and of the consonants between two vowels only
/// the last starts the next syllable (ka-lib-ras-yon, kont-rol).
/// </summary>
internal static class DocxHyphenation
{
    private const string TurkishVowels = "aâeıiîoöuûüAÂEIİÎOÖUÛÜ";

    /// <summary>
    /// Whether a language (such as <c>tr-TR</c>) has rules here.
    /// </summary>
    public static bool Knows(string? language) =>
        language != null && (language.Equals("tr", StringComparison.OrdinalIgnoreCase) || language.StartsWith("tr-", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// The positions in a word where it may be broken, in order: the break comes before the character at each
    /// position. At least two letters stay on each side; punctuation after the word stays with its end, and words with
    /// other characters inside are not broken.
    /// </summary>
    public static List<int> Points(string word)
    {
        var points = new List<int>();
        var length = word.Length;
        while (length > 0 && !char.IsLetter(word[length - 1]))
        {
            length--;
        }

        if (length < 4)
        {
            return points;
        }

        for (var i = 0; i < length; i++)
        {
            if (!char.IsLetter(word[i]))
            {
                return points;
            }
        }

        var previous = -1;
        for (var i = 0; i < length; i++)
        {
            if (TurkishVowels.IndexOf(word[i]) < 0)
            {
                continue;
            }

            if (previous >= 0)
            {
                // No consonant between the vowels: the break falls between them; else before the last consonant.
                var point = i - previous == 1 ? i : i - 1;
                if (point >= 2 && length - point >= 2)
                {
                    points.Add(point);
                }
            }

            previous = i;
        }

        return points;
    }
}
