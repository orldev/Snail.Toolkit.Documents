using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents;

/// <summary>Reads a file by its name: picks the reader its extension calls for and holds the reading to its limits.</summary>
/// <remarks>
/// The format is told by the extension alone. Sniffing the content would read a file as something its sender did not
/// call it, and a file that is not what its name says is answered as damaged, which is what it is.
/// </remarks>
public sealed class DocumentReader(IEnumerable<IFormatReader> readers, TimeProvider? clock = null)
{
    private readonly IReadOnlyList<IFormatReader> _readers = [.. readers];

    private readonly TimeProvider _clock = clock ?? TimeProvider.System;

    /// <summary>A reader of every format this library reads.</summary>
    public static DocumentReader Every { get; } = new(FormatReaders.Every());

    /// <summary>The extensions read, with their dots, in lower case.</summary>
    public IReadOnlyList<string> Extensions { get; } = [.. readers.SelectMany(reader => reader.Extensions)];

    /// <summary>The reader a file of this name is read with; null when no reader takes it.</summary>
    public IFormatReader? ReaderFor(string fileName)
    {
        string extension = Path.GetExtension(fileName).ToLowerInvariant();

        return _readers.FirstOrDefault(reader => reader.Extensions.Contains(extension, StringComparer.Ordinal));
    }

    /// <summary>The file read into its document, or why it was not, within <see cref="ReadingLimits.MostReadingSeconds"/>.</summary>
    /// <remarks>
    /// The format's reader runs on the pool under the deadline, so the caller holds no thread while a file is parsed and gets
    /// its answer at the deadline whatever the reader is doing: a reader past it is left to stop at its next check of the
    /// token. A format's library throws what it likes on a damaged file — a bare exception, an argument out of range, a
    /// state it did not expect — and every such throw is the file's fault, so it is answered as damaged with the library's
    /// words. Only the caller's cancellation and the machine's memory are not the file's: those propagate.
    /// </remarks>
    public async Task<Reading> ReadAsync(Stream content, string fileName, CancellationToken cancellationToken = default)
    {
        if (ReaderFor(fileName) is not { } reader)
            return Reading.Refused(ReadingFailure.UnknownFormat, $"'{Path.GetExtension(fileName)}' is not a format this library reads");

        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(ReadingLimits.MostReadingSeconds), _clock);
        using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);

        try
        {
            var blocks = await Task.Run(() => reader.ReadAsync(content, cancellation.Token), cancellation.Token).WaitAsync(cancellation.Token);
            var document = Document.Of(reader.Kind, blocks);

            return Reading.Of(document);
        }
        catch (OperationCanceledException) when (deadline.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
        {
            return Reading.Refused(ReadingFailure.TooSlow, $"reading took longer than {ReadingLimits.MostReadingSeconds} seconds");
        }
        catch (DocumentTooLargeException failure)
        {
            return Reading.Refused(ReadingFailure.TooLarge, failure.Message);
        }
        catch (DocumentLockedException failure)
        {
            return Reading.Refused(ReadingFailure.Locked, failure.Message);
        }
        catch (Exception failure) when (failure is not OperationCanceledException and not OutOfMemoryException)
        {
            return Reading.Refused(ReadingFailure.Damaged, failure.Message);
        }
    }
}
