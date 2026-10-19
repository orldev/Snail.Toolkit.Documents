using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Exceptions;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>A PDF through its text layer: every page a part, its text in reading order.</summary>
/// <remarks>
/// A page that carries pictures and no text is a scan. Its words are not in the file as text, so it is marked as
/// one rather than passed over as an empty page: the model is told that the page has something on it and that
/// reading it takes a model that sees pictures.
/// </remarks>
public sealed class PdfReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Pdf;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".pdf"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        using var pdf = Opened(content);
        List<Block> blocks = [];

        foreach (var page in pdf.GetPages())
        {
            cancellationToken.ThrowIfCancellationRequested();

            string text = ContentOrderTextExtractor.GetText(page).ReplaceLineEndings("\n");
            bool isScanned = string.IsNullOrWhiteSpace(text) && page.GetImages().Any();

            blocks.Add(new PartStart(PartKind.Page, page.Number, null, isScanned));
            blocks.AddRange(text
                .Split("\n\n", StringSplitOptions.RemoveEmptyEntries)
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph))
                .Select(paragraph => new Paragraph(paragraph.Trim('\n'))));
        }

        return blocks;
    }

    /// <summary>The PDF opened, every way it can fail to open turned into the failure the readers share.</summary>
    internal static PdfDocument Opened(Stream content)
    {
        try
        {
            return PdfDocument.Open(content);
        }
        catch (PdfDocumentEncryptedException failure)
        {
            throw new DocumentLockedException("the PDF is protected by a password", failure);
        }
        catch (Exception failure) when (failure is PdfDocumentFormatException or PdfDocumentStackDepthException)
        {
            throw new InvalidDataException(failure.Message, failure);
        }
    }
}
