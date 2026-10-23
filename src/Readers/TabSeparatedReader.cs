namespace Snail.Toolkit.Documents.Readers;

/// <summary>Values separated by tabs, a .tsv: one table, its first line the header.</summary>
/// <remarks>
/// The name says the separator, so none is guessed: guessed from the first line, a header of «Name, first, last» and
/// one tab was split by its commas into three columns where the file has two.
/// </remarks>
public sealed class TabSeparatedReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Csv;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".tsv"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(DelimitedReader.Read(content, '\t', cancellationToken));
}
