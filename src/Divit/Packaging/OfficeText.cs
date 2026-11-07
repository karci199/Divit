#nullable enable
namespace Divit.Packaging;

using Divit.Properties;
using System;
using System.Globalization;
using System.Text;

/// <summary>
/// Text written into Office XML: culture-independent numbers, colours, characters XML can hold. The buffer methods
/// format without allocating a string, for values written millions of times (cell addresses, cell values).
/// </summary>
internal static class OfficeText
{
    /// <summary>
    /// A whole number in invariant form.
    /// </summary>
    public static string Int(int value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Drops control characters that XML cannot hold (tab, line feed and carriage return stay); XmlWriter throws on
    /// them. Returns the same string when there is nothing to drop.
    /// </summary>
    public static string Clean(string text)
    {
        StringBuilder? sb = null;
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            var bad = ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r';
            if (bad && sb == null)
            {
                sb = new StringBuilder(text, 0, i, text.Length);
            }
            else if (!bad)
            {
                sb?.Append(ch);
            }
        }

        return sb?.ToString() ?? text;
    }

    /// <summary>
    /// <c>#RRGGBB</c> (with or without <c>#</c>, any case) → <c>RRGGBB</c> in upper case, as DrawingML and
    /// WordprocessingML write colours.
    /// </summary>
    /// <exception cref="ArgumentException">The text is not six hex digits.</exception>
    public static string Rgb(string color)
    {
        var hex = color.TrimStart('#').ToUpperInvariant();
        if (hex.Length != 6)
        {
            throw new ArgumentException(Messages.Format(Strings.ColorInvalid, color));
        }

        foreach (var ch in hex)
        {
            if (!((ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'F')))
            {
                throw new ArgumentException(Messages.Format(Strings.ColorInvalid, color));
            }
        }

        return hex;
    }

    /// <summary>
    /// Whether the bytes start with the PNG signature.
    /// </summary>
    public static bool IsPng(byte[] data) =>
        data.Length >= 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
        && data[4] == 0x0D && data[5] == 0x0A && data[6] == 0x1A && data[7] == 0x0A;

    /// <summary>
    /// Writes a non-negative whole number into the buffer from <paramref name="position"/>; returns the end position.
    /// </summary>
    public static int FormatInt(char[] buffer, int position, int value)
    {
        var digits = 1;
        for (var rest = value; rest >= 10; rest /= 10)
        {
            digits++;
        }

        var end = position + digits;
        for (var i = end - 1; i >= position; i--)
        {
            buffer[i] = (char)('0' + (value % 10));
            value /= 10;
        }

        return end;
    }

    /// <summary>
    /// Writes a number into the buffer (at least 32 characters) as <c>ToString("R")</c> would; returns its length.
    /// </summary>
    /// <remarks>
    /// Whole numbers below 10^15 (counters, years, IDs) are written digit by digit on every runtime: "R" gives them
    /// without exponent or point, so the text is the same. Other numbers go to the runtime's own round-trip format.
    /// </remarks>
    public static int FormatDouble(char[] buffer, double number)
    {
        // Negative zero is left to the runtime: .NET writes "-0", .NET Framework "0".
        if (number == Math.Floor(number) && Math.Abs(number) < 1e15 && (number != 0 || BitConverter.DoubleToInt64Bits(number) == 0))
        {
            var whole = (long)number;
            var position = 0;
            if (whole < 0)
            {
                buffer[position++] = '-';
                whole = -whole;
            }

            var digits = 1;
            for (var rest = whole; rest >= 10; rest /= 10)
            {
                digits++;
            }

            var end = position + digits;
            for (var i = end - 1; i >= position; i--)
            {
                buffer[i] = (char)('0' + (int)(whole % 10));
                whole /= 10;
            }

            return end;
        }

#if NET
        number.TryFormat(buffer, out var length, "R", CultureInfo.InvariantCulture);
        return length;
#else
        var text = number.ToString("R", CultureInfo.InvariantCulture);
        text.CopyTo(0, buffer, 0, text.Length);
        return text.Length;
#endif
    }
}
