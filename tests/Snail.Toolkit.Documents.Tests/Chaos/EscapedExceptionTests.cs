using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace Snail.Toolkit.Documents.Tests.Chaos;

/// <summary>The README promises that <c>ReadAsync</c> throws only when its caller cancels it; these are the files that made it throw before every failure of a format's library was answered as damaged.</summary>
public sealed class EscapedExceptionTests
{
    /// <summary>A workbook whose sheet list names a relationship the package does not have: the SDK answered with an argument out of range; the reader asks first.</summary>
    [Fact]
    public async Task Read_WorkbookWhoseSheetNamesAPartItDoesNotHave_IsARefusalNotAThrow()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(WorkbookWithADanglingSheet()), "dangling.xlsx");

        Assert.Equal(ReadingFailure.Damaged, reading.Failure);
    }

    /// <summary>A presentation whose slide list names a relationship the package does not have.</summary>
    [Fact]
    public async Task Read_PresentationWhoseSlideNamesAPartItDoesNotHave_IsARefusalNotAThrow()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(DeckWithADanglingSlide()), "dangling.pptx");

        Assert.Equal(ReadingFailure.Damaged, reading.Failure);
    }

    /// <summary>Only the Office readers copy a stream that cannot seek into memory; a PDF from a pipe reaches PdfPig as it is.</summary>
    [Fact]
    public async Task Read_PdfFromAStreamThatCannotSeek_IsRead()
    {
        var reading = await DocumentReader.Every.ReadAsync(new ForwardOnlyStream(Samples.Pdf()), "piped.pdf");

        Assert.True(reading.IsRead, reading.Reason);
    }

    /// <summary>An older workbook from a pipe reaches ExcelDataReader as it is.</summary>
    [Fact]
    public async Task Read_XlsFromAStreamThatCannotSeek_IsRead()
    {
        var reading = await DocumentReader.Every.ReadAsync(new ForwardOnlyStream(Samples.OldExcel()), "piped.xls");

        Assert.True(reading.IsRead, reading.Reason);
    }

    /// <summary>A record that names a key twice is a JSON object System.Text.Json builds lazily, and the duplicate threw when the object was first walked; refused as it is parsed now.</summary>
    [Fact]
    public async Task Read_JsonRecordWithADuplicateKey_IsARefusalOrADocumentNotAThrow()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream("""[{"a":1,"a":2}]"""u8.ToArray()), "twice.json");

        Assert.True(reading.IsRead || reading.Failure == ReadingFailure.Damaged, reading.Reason);
    }

    /// <summary>
    /// Every file of a format, cut short, with bytes flipped or with a window zeroed, read through the reader of its name: whatever the
    /// format's library throws on it must come back as a reading. An exception that escapes is a crash in every program that reads
    /// documents with this library.
    /// </summary>
    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    [InlineData("pdf")]
    [InlineData("xls")]
    public async Task Read_MutatedFile_IsAlwaysARefusalOrADocument(string extension)
    {
        byte[] original = Sample(extension);
        var chaos = new Random(20261021);
        List<string> escaped = [];

        for (int round = 0; round < 150; round++)
        {
            byte[] mutated = Mutated(original, chaos, round % 3);

            try
            {
                await DocumentReader.Every.ReadAsync(new MemoryStream(mutated), $"mutated.{extension}");
            }
            catch (Exception failure) when (failure is not OperationCanceledException)
            {
                escaped.Add($"round {round} ({Kind(round % 3)}): {failure.GetType().Name}: {failure.Message}");
            }
        }

        Assert.True(escaped.Count == 0, $"{escaped.Count} of 150 mutations escaped:\n{string.Join('\n', escaped.Take(8))}");
    }

    private static byte[] Sample(string extension) => extension switch
    {
        "docx" => Samples.Word(),
        "xlsx" => Samples.ExcelOfEveryCell(),
        "pptx" => Samples.Slides(),
        "pdf" => Samples.Pdf(),
        _ => Samples.OldExcel()
    };

    private static string Kind(int mutation) => mutation switch
    {
        0 => "truncated",
        1 => "bytes flipped",
        _ => "window zeroed"
    };

    private static byte[] Mutated(byte[] original, Random chaos, int mutation)
    {
        if (mutation == 0)
            return original[..chaos.Next(0, original.Length)];

        byte[] copy = [.. original];

        if (mutation == 1)
        {
            for (int flip = 0; flip < 8; flip++)
                copy[chaos.Next(copy.Length)] ^= 0xFF;

            return copy;
        }

        int window = chaos.Next(0, Math.Max(1, copy.Length - 64));
        Array.Clear(copy, window, Math.Min(64, copy.Length - window));

        return copy;
    }

    private static byte[] WorkbookWithADanglingSheet()
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook(new S.Sheets(new S.Sheet { Id = "rId404", SheetId = 1, Name = "Gone" }));
        }

        return stream.ToArray();
    }

    private static byte[] DeckWithADanglingSlide()
    {
        using var stream = new MemoryStream();

        using (var deck = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentation = deck.AddPresentationPart();
            presentation.Presentation = new P.Presentation(new P.SlideIdList(new P.SlideId { Id = 256, RelationshipId = "rId404" }));
        }

        return stream.ToArray();
    }
}
