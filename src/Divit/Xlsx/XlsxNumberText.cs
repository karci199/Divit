#nullable enable
namespace Divit.Xlsx;

using System;
using System.Globalization;
using System.Text;

/// <summary>
/// A cell value as Excel shows it with its number format: General, numbers (<c>#,##0.00</c>, <c>0%</c>,
/// <c>0.00E+00</c>), sections for negative and zero values, dates and times (<c>yyyy-mm-dd hh:mm</c>), text
/// (<c>@</c>), quoted and escaped literals. Colours, conditions, padding and fill characters are left out.
/// </summary>
internal static class XlsxNumberText
{
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    public static string Format(object? value, string? format)
    {
        switch (value)
        {
            case null:
                return string.Empty;
            case bool b:
                return b ? "TRUE" : "FALSE";
            case string s:
                return Section(format, 3) is string text && text.Contains("@") ? Literal(text).Replace("@", s) : s;
            case XlsxRichText rich:
                return rich.Text;
            case double d:
                return Number(d, format);
            default:
                return Convert.ToString(value, Invariant) ?? string.Empty;
        }
    }

    private static string Number(double value, string? format)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return "#NUM!";
        }

        if (string.IsNullOrEmpty(format) || format == "General" || format == "@")
        {
            return General(value);
        }

        // Sections: positive; negative; zero. The negative section shows the value without its sign.
        var sections = Split(format!);
        var section = sections[0];
        var abs = false;
        if (value < 0 && sections.Length > 1 && sections[1].Length > 0)
        {
            section = sections[1];
            abs = true;
        }
        else if (value == 0 && sections.Length > 2 && sections[2].Length > 0)
        {
            section = sections[2];
        }

        var clean = Clean(section);
        if (clean.Trim() == "General")
        {
            return General(abs ? Math.Abs(value) : value);
        }

        if (IsDate(clean))
        {
            return Date(value, clean);
        }

        var shown = abs ? Math.Abs(value) : value;
        var text = shown.ToString(clean, Invariant);

        // A value that rounds to zero shows no minus sign, as in Excel.
        if (text.StartsWith("-", StringComparison.Ordinal) && text.IndexOfAny("123456789".ToCharArray()) < 0)
        {
            text = text.Substring(1);
        }

        return text;
    }

    /// <summary>
    /// General: up to ten significant digits, without trailing zeros; scientific for very large or small values.
    /// </summary>
    private static string General(double value)
    {
        if (value == 0)
        {
            return "0";
        }

        var magnitude = Math.Abs(value);
        if (magnitude >= 1e11 || magnitude < 1e-9)
        {
            return value.ToString("0.#####E+00", Invariant);
        }

        var digits = Math.Max(0, 10 - (int)Math.Floor(Math.Log10(magnitude)) - 1);
        return Math.Round(value, Math.Min(15, digits)).ToString("0.##########", Invariant);
    }

    private static string? Section(string? format, int index)
    {
        if (format == null)
        {
            return null;
        }

        var sections = Split(format);
        return index < sections.Length ? sections[index] : sections.Length == 1 ? sections[0] : null;
    }

    /// <summary>
    /// Splits sections at semicolons outside quotes.
    /// </summary>
    private static string[] Split(string format)
    {
        var parts = new System.Collections.Generic.List<string>();
        var current = new StringBuilder();
        var quoted = false;
        for (var i = 0; i < format.Length; i++)
        {
            var c = format[i];
            if (c == '"') { quoted = !quoted; }
            if (c == '\\' && i + 1 < format.Length) { current.Append(c).Append(format[++i]); continue; }
            if (c == ';' && !quoted) { parts.Add(current.ToString()); current.Clear(); continue; }
            current.Append(c);
        }

        parts.Add(current.ToString());
        return parts.ToArray();
    }

    /// <summary>
    /// A section without what only changes looks: [colours] and [conditions], padding (<c>_x</c>) as a space, fill
    /// characters (<c>*x</c>) dropped.
    /// </summary>
    private static string Clean(string section)
    {
        var text = new StringBuilder();
        var quoted = false;
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            if (c == '"') { quoted = !quoted; text.Append(c); continue; }
            if (quoted) { text.Append(c); continue; }
            if (c == '[')
            {
                var end = section.IndexOf(']', i);
                if (end < 0) { break; }

                // Elapsed time ([h]) stays; colours, conditions and locales go.
                var inner = section.Substring(i + 1, end - i - 1);
                if (inner.Length > 0 && (inner[0] == 'h' || inner[0] == 'm' || inner[0] == 's')) { text.Append(inner); }
                i = end;
                continue;
            }

            if (c == '_' && i + 1 < section.Length) { text.Append(' '); i++; continue; }
            if (c == '*' && i + 1 < section.Length) { i++; continue; }
            text.Append(c);
        }

        return text.ToString();
    }

    private static string Literal(string section)
    {
        var text = new StringBuilder();
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            if (c == '"') { continue; }
            if (c == '\\' && i + 1 < section.Length) { text.Append(section[++i]); continue; }
            text.Append(c);
        }

        return text.ToString();
    }

    private static bool IsDate(string section)
    {
        var quoted = false;
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            if (c == '"') { quoted = !quoted; continue; }
            if (quoted) { continue; }
            if (c == '\\') { i++; continue; }
            if ("yYdDhHsS".IndexOf(c) >= 0 || ((c == 'm' || c == 'M') && section.IndexOf('0') < 0 && section.IndexOf('#') < 0))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// A date and time from Excel's serial number: m is the month unless it follows hours or comes before seconds,
    /// where it is minutes.
    /// </summary>
    private static string Date(double serial, string section)
    {
        DateTime date;
        try
        {
            date = DateTime.FromOADate(serial);
        }
        catch (ArgumentException)
        {
            return "#####";
        }

        var ampm = section.IndexOf("AM/PM", StringComparison.OrdinalIgnoreCase) >= 0;
        var text = new StringBuilder();
        var lastWasHour = false;
        for (var i = 0; i < section.Length;)
        {
            var c = section[i];
            if (c == '"')
            {
                var end = section.IndexOf('"', i + 1);
                if (end < 0) { end = section.Length; }
                text.Append(section, i + 1, Math.Max(0, end - i - 1));
                i = end + 1;
                continue;
            }

            if (c == '\\' && i + 1 < section.Length)
            {
                text.Append(section[i + 1]);
                i += 2;
                continue;
            }

            if (string.Compare(section, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                text.Append(date.Hour < 12 ? "AM" : "PM");
                i += 5;
                continue;
            }

            var lower = char.ToLowerInvariant(c);
            var run = 1;
            while (i + run < section.Length && char.ToLowerInvariant(section[i + run]) == lower) { run++; }
            switch (lower)
            {
                case 'y':
                    text.Append(run <= 2 ? (date.Year % 100).ToString("00", Invariant) : date.Year.ToString("0000", Invariant));
                    break;
                case 'd':
                    text.Append(run == 1 ? date.Day.ToString(Invariant) : run == 2 ? date.Day.ToString("00", Invariant)
                        : run == 3 ? date.ToString("ddd", Invariant) : date.ToString("dddd", Invariant));
                    break;
                case 'h':
                    var hour = ampm ? ((date.Hour + 11) % 12) + 1 : date.Hour;
                    text.Append(run == 1 ? hour.ToString(Invariant) : hour.ToString("00", Invariant));
                    lastWasHour = true;
                    i += run;
                    continue;
                case 's':
                    text.Append(run == 1 ? date.Second.ToString(Invariant) : date.Second.ToString("00", Invariant));
                    break;
                case 'm':
                    var minutes = lastWasHour || NextIsSeconds(section, i + run);
                    if (minutes)
                    {
                        text.Append(run == 1 ? date.Minute.ToString(Invariant) : date.Minute.ToString("00", Invariant));
                    }
                    else
                    {
                        text.Append(run == 1 ? date.Month.ToString(Invariant) : run == 2 ? date.Month.ToString("00", Invariant)
                            : run == 3 ? date.ToString("MMM", Invariant) : run == 4 ? date.ToString("MMMM", Invariant) : date.ToString("MMMM", Invariant).Substring(0, 1));
                    }

                    break;
                default:
                    text.Append(section, i, run);
                    break;
            }

            if (lower != ' ' && lower != ':' && lower != '.') { lastWasHour = false; }
            i += run;
        }

        return text.ToString();
    }

    private static bool NextIsSeconds(string section, int from)
    {
        for (var i = from; i < section.Length; i++)
        {
            var c = char.ToLowerInvariant(section[i]);
            if (c == 's') { return true; }
            if (char.IsLetter(c)) { return false; }
        }

        return false;
    }
}
