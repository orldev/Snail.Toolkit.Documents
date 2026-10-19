namespace Snail.Toolkit.Documents.Readers;

/// <summary>The readers of every format this library reads.</summary>
public static class Formats
{
    /// <summary>A reader of each format, new; none of them keeps anything between files.</summary>
    public static IReadOnlyList<IFormatReader> Every() =>
    [
        new WordReader(),
        new ExcelReader(),
        new PowerPointReader(),
        new PdfReader(),
        new HtmlReader(),
        new DelimitedReader(),
        new JsonReader(),
        new MarkdownReader(),
        new PlainTextReader()
    ];
}
