using System.Xml;
using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents;

/// <summary>Reads a file by its name: picks the reader its extension calls for and holds the reading to its limits.</summary>
/// <remarks>
/// The format is told by the extension alone. Sniffing the content would read a file as something its sender did not
/// call it, and a file that is not what its name says is answered as damaged, which is what it is.
/// </remarks>
public sealed class DocumentReader(IEnumerable<IFormatReader> formats, TimeProvider? clock = null)
{
    private readonly IReadOnlyList<IFormatReader> _formats = [.. formats];

    private readonly TimeProvider _clock = clock ?? TimeProvider.System;

    /// <summary>A reader of every format this library reads.</summary>
    public static DocumentReader Every { get; } = new(Formats.Every());

    /// <summary>The extensions read, with their dots, in lower case.</summary>
    public IReadOnlyList<string> Extensions => [.. _formats.SelectMany(format => format.Extensions)];

    /// <summary>The reader a file of this name is read with; null when no reader takes it.</summary>
    public IFormatReader? ReaderFor(string fileName)
    {
        string extension = Path.GetExtension(fileName).ToLowerInvariant();

        return _formats.FirstOrDefault(format => format.Extensions.Contains(extension, StringComparer.Ordinal));
    }

    /// <summary>The file read into its document, or why it was not, within <see cref="ReadingLimits.MostReadingSeconds"/>.</summary>
    /// <remarks>
    /// A file built to keep a parser busy costs half a minute of one thread and no more: the deadline stops it and the
    /// answer says it was too slow. Cancelled by the caller, the reading throws as any cancelled call does — that is the
    /// caller's decision, not a fault of the file.
    /// </remarks>
    public async Task<Reading> ReadAsync(Stream content, string fileName, CancellationToken cancellationToken = default)
    {
        if (ReaderFor(fileName) is not { } format)
            return Reading.Refused(ReadingFailure.UnknownFormat, $"'{Path.GetExtension(fileName)}' is not a format this library reads");

        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(ReadingLimits.MostReadingSeconds), _clock);
        using var reading = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);

        try
        {
            var blocks = await format.ReadAsync(content, reading.Token).WaitAsync(reading.Token);

            return Reading.Of(Markdown.Of(format.Format, blocks));
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
        catch (Exception failure) when (failure is InvalidDataException or XmlException)
        {
            return Reading.Refused(ReadingFailure.Damaged, failure.Message);
        }
    }
}
