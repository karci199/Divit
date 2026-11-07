namespace Divit.Properties;

using System.Globalization;
using System.Resources;

/// <summary>
/// Keys of the messages in <c>Strings.resx</c>; one constant per resource name.
/// </summary>
/// <remarks>
/// Written by hand instead of a generated designer file: the SDK project has no designer step, and constants keep the
/// keys checked by the compiler.
/// </remarks>
internal static class Strings
{
    internal const string SeriesRangeEmpty = nameof(SeriesRangeEmpty);
    internal const string ImageNotPng = nameof(ImageNotPng);
    internal const string MergeRangeReversed = nameof(MergeRangeReversed);
    internal const string CellOutOfRange = nameof(CellOutOfRange);
    internal const string SheetNameLength = nameof(SheetNameLength);
    internal const string SheetNameCharacters = nameof(SheetNameCharacters);
    internal const string SheetNameDuplicate = nameof(SheetNameDuplicate);
    internal const string WorkbookEmpty = nameof(WorkbookEmpty);
    internal const string ColorInvalid = nameof(ColorInvalid);
    internal const string DateOutOfRange = nameof(DateOutOfRange);
    internal const string LinkEmpty = nameof(LinkEmpty);
    internal const string GroupRangeReversed = nameof(GroupRangeReversed);
    internal const string NoVisibleSheet = nameof(NoVisibleSheet);
    internal const string ActiveSheetHidden = nameof(ActiveSheetHidden);
    internal const string SheetNotInWorkbook = nameof(SheetNotInWorkbook);
    internal const string DefinedNameInvalid = nameof(DefinedNameInvalid);
    internal const string DefinedNameDuplicate = nameof(DefinedNameDuplicate);
}

/// <summary>
/// Reads messages from the resources in the user's UI language; English is the neutral language.
/// </summary>
internal static class Messages
{
    private static readonly ResourceManager Resources = new ResourceManager("Divit.Properties.Strings", typeof(Messages).Assembly);

    /// <summary>
    /// The message for a key.
    /// </summary>
    /// <param name="key">One of the <see cref="Strings"/> constants.</param>
    /// <returns>The message; the key itself when it is missing.</returns>
    internal static string Get(string key) => Resources.GetString(key, CultureInfo.CurrentUICulture) ?? key;

    /// <summary>
    /// The message for a key with its placeholders filled.
    /// </summary>
    /// <param name="key">One of the <see cref="Strings"/> constants.</param>
    /// <param name="args">Values for the placeholders.</param>
    /// <returns>The formatted message.</returns>
    internal static string Format(string key, params object[] args) => string.Format(CultureInfo.CurrentCulture, Get(key), args);
}
