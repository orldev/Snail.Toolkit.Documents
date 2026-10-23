namespace Snail.Toolkit.Documents.Readers;

/// <summary>The readers of every format this library reads.</summary>
/// <remarks>A reader is named after the <see cref="DocumentKind"/> it reads; a second reader of the same kind is named after its extension, as <see cref="XlsReader"/> and <see cref="TsvReader"/> are.</remarks>
public static class FormatReaders
{
    /// <summary>A reader of each format, new; none of them keeps anything between files.</summary>
    public static IReadOnlyList<IFormatReader> Every() =>
    [
        new WordReader(),
        new ExcelReader(),
        new XlsReader(),
        new PowerPointReader(),
        new PdfReader(),
        new HtmlReader(),
        new CsvReader(),
        new TsvReader(),
        new JsonReader(),
        new MarkdownReader(),
        new PlainTextReader()
    ];
}
