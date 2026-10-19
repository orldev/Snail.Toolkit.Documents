using System.IO.Compression;
using System.Text;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Time.Testing;

namespace Snail.Toolkit.Documents.Tests;

/// <summary>A file read by its name: the reader its extension calls for, and every way it can fail answered as a reading that says why.</summary>
public sealed class DocumentReaderTests
{
    [Fact]
    public async Task Read_PdfNamedInCapitals_IsReadAsAPdf()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(Samples.Pdf()), "REPORT.PDF");

        Assert.True(reading.IsRead);
        Assert.Equal(DocumentKind.Pdf, reading.Document!.Format);
    }

    /// <summary>An older binary Word file has a name no reader takes: refused before a byte is read.</summary>
    [Fact]
    public async Task Read_NameNoReaderTakes_IsRefusedAsAnUnknownFormat()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream([1, 2, 3]), "letter.rtf");

        Assert.Equal(ReadingFailure.UnknownFormat, reading.Failure);
        Assert.Contains(".rtf", reading.Reason, StringComparison.Ordinal);
    }

    /// <summary>A valid package around a broken part fails only when the part is first read, as an <c>XmlException</c>; it is still a damaged file.</summary>
    [Fact]
    public async Task Read_WordWithBrokenDocumentXml_IsRefusedAsDamaged()
    {
        byte[] broken = WithEntry(Samples.Word(), "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>");

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(broken), "broken.docx");

        Assert.Equal(ReadingFailure.Damaged, reading.Failure);
    }

    [Fact]
    public async Task Read_BytesThatAreNotThePdfTheNameSays_IsRefusedAsDamaged()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(Encoding.UTF8.GetBytes("plain words")), "scan.pdf");

        Assert.Equal(ReadingFailure.Damaged, reading.Failure);
    }

    [Fact]
    public async Task Read_PdfUnderAPassword_IsRefusedAsLocked()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(Samples.LockedPdf()), "locked.pdf");

        Assert.Equal(ReadingFailure.Locked, reading.Failure);
    }

    [Fact]
    public async Task Read_JsonThatUnpacksPastTheCells_IsRefusedAsTooLarge()
    {
        string records = $"[{string.Join(',', Enumerable.Range(0, 2_000).Select(index => $$"""{"k{{index}}":1}"""))}]";

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(Encoding.UTF8.GetBytes(records)), "sparse.json");

        Assert.Equal(ReadingFailure.TooLarge, reading.Failure);
    }

    /// <summary>A reader still busy when the half minute is up is stopped, and the answer says it was too slow rather than throwing.</summary>
    [Fact]
    public async Task Read_ReaderPastTheDeadline_IsRefusedAsTooSlow()
    {
        var clock = new FakeTimeProvider();
        var reader = new DocumentReader([new Endless()], clock);

        var reading = reader.ReadAsync(new MemoryStream(), "endless.txt");
        clock.Advance(TimeSpan.FromSeconds(ReadingLimits.MostReadingSeconds));

        Assert.Equal(ReadingFailure.TooSlow, (await reading).Failure);
    }

    [Fact]
    public async Task Read_CancelledByTheCaller_Throws()
    {
        using var cancelled = new CancellationTokenSource();
        await cancelled.CancelAsync();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DocumentReader.Every.ReadAsync(new MemoryStream(Samples.Csv()), "rows.csv", cancelled.Token));
    }

    [Fact]
    public void AddDocumentReading_Container_HoldsAReaderOfEveryFormat()
    {
        using var provider = new ServiceCollection().AddDocumentReading().BuildServiceProvider();

        var reader = provider.GetRequiredService<DocumentReader>();

        Assert.Equal(Documents.Readers.Formats.Every().Count, provider.GetServices<IFormatReader>().Count());
        Assert.Contains(".pptx", reader.Extensions);
        Assert.Contains(".xls", reader.Extensions);
    }

    private static byte[] WithEntry(byte[] package, string entry, string text)
    {
        using var stream = new MemoryStream();
        stream.Write(package);

        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry(entry)!.Delete();

            using var writer = new StreamWriter(zip.CreateEntry(entry).Open());
            writer.Write(text);
        }

        return stream.ToArray();
    }

    /// <summary>A reader that never finishes until it is told to stop.</summary>
    private sealed class Endless : IFormatReader
    {
        public DocumentKind Format => DocumentKind.Text;

        public IReadOnlyList<string> Extensions { get; } = [".txt"];

        public async Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default)
        {
            await Task.Delay(Timeout.Infinite, cancellationToken);

            return [];
        }
    }
}
