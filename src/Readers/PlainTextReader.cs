namespace Snail.Toolkit.Documents.Readers;

/// <summary>Plain text: paragraphs where the file has blank lines between them.</summary>
public sealed class PlainTextReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Kind => DocumentKind.Text;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".txt", ".text", ".log"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content));

    private static IReadOnlyList<Block> Read(Stream content) =>
    [
        .. Decoding.Text(content)
            .ReplaceLineEndings("\n")
            .Split("\n\n", StringSplitOptions.RemoveEmptyEntries)
            .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph))
            .Select(paragraph => new Paragraph(paragraph.Trim('\n')))
    ];
}
