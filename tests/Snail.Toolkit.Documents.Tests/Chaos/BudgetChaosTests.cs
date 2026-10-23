using System.Diagnostics;
using Snail.Toolkit.Documents.Readers;
using Xunit.Abstractions;

namespace Snail.Toolkit.Documents.Tests.Chaos;

/// <summary>The budgets bound what a document becomes and what reading it costs: a file past them is refused for less than its own size.</summary>
/// <remarks>Run apart from every other test: the peak of the heap is the process's, and a neighbour's fifty megabytes would be counted against the reader.</remarks>
[Collection(PeakHeapCollection.Name)]
public sealed class BudgetChaosTests(ITestOutputHelper output)
{
    /// <summary>
    /// Decoded whole, its line endings replaced in a second copy and split into paragraphs in a third, a text file of fifty megabytes cost
    /// 809 MB to refuse. The bytes of a file bound its characters from below in every encoding the library decodes, so a file past twice
    /// the budget is refused unread. The reader is called on this thread so that only its allocations are measured.
    /// </summary>
    [Fact]
    public async Task Read_TextFileTwiceTheCharacterBudget_IsRefusedForLessThanItsOwnSize()
    {
        byte[] prose = Hostility.Prose(50);
        long before = GC.GetAllocatedBytesForCurrentThread();

        await Assert.ThrowsAsync<DocumentTooLargeException>(() => new PlainTextReader().ReadAsync(new MemoryStream(prose)));

        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(allocated < prose.Length, $"refusing {prose.Length / 1_048_576} MB of text allocated {allocated / 1_048_576} MB");
    }

    /// <summary>
    /// The zip limits let a worksheet of forty megabytes through. Loaded as a tree before the first row was read, it stood at 1.5 GB for
    /// 16 s until the budget of cells refused it — one step the token never interrupted, in a container of 512 MB that the budget of
    /// characters was sized for. Read row by row, the heap holds the rows the budget allows and little else.
    /// </summary>
    [Fact]
    public async Task Read_WorksheetOfAMillionRowsStoredUncompressed_IsRefusedWithinTheMemoryTheLimitsPromise()
    {
        byte[] workbook = Hostility.TallWorkbook(1_050_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(workbook), "tall.xlsx");

        long peak = heap.Bytes;
        output.WriteLine($"refusing a workbook of {workbook.Length / 1_048_576} MB took {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.Equal(ReadingFailure.TooLarge, reading.Failure);
        Assert.True(peak < 512L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>
    /// The zip limits let a document of sixty megabytes through, and the budget of characters refuses it at the twenty-millionth
    /// character; what the heap holds until then is the question. Loaded as a tree, the whole body stood in memory before the first
    /// paragraph was read; read paragraph by paragraph, the heap holds the blocks the budget allows.
    /// </summary>
    [Fact]
    public async Task Read_DocumentOfAMillionParagraphsStoredUncompressed_IsRefusedWithinTheMemoryTheLimitsPromise()
    {
        byte[] document = Hostility.TallDocument(1_000_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(document), "tall.docx");

        long peak = heap.Bytes;
        output.WriteLine($"refusing a document of {document.Length / 1_048_576} MB took {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.Equal(ReadingFailure.TooLarge, reading.Failure);
        Assert.True(peak < 512L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>
    /// One paragraph of a thousand runs of fifty thousand letters is fifty million characters, past the budget: built as a tree, the
    /// paragraph stood whole before its first letter was counted. Read run by run under the budget, the heap holds the budget and the file.
    /// </summary>
    [Fact]
    public async Task Read_ParagraphOfFiftyMillionLetters_IsRefusedWithinTheMemoryTheLimitsPromise()
    {
        byte[] document = Hostility.WideParagraph(1_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(document), "wide.docx");

        long peak = heap.Bytes;
        output.WriteLine($"refusing a paragraph of {document.Length / 1_048_576} MB took {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.Equal(ReadingFailure.TooLarge, reading.Failure);
        Assert.True(peak < 256L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>
    /// A table of nine hundred thousand short rows is under every budget and must be read: built as a tree, the table stood whole in
    /// memory; read cell by cell, the heap holds the rows as strings.
    /// </summary>
    [Fact]
    public async Task Read_TableOfNineHundredThousandRows_IsReadWithinTheMemoryTheLimitsPromise()
    {
        byte[] document = Hostility.TallTable(900_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(document), "table.docx");

        long peak = heap.Bytes;
        output.WriteLine($"reading a table of {document.Length / 1_048_576} MB took {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.True(reading.IsRead, reading.Reason);
        Assert.Equal(900_000, reading.Document!.Tables[0].Rows.Count);
        Assert.True(peak < 512L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>One shape of two thousand paragraphs of fifty thousand letters: built as a tree, the shape stood whole; read under the budget, it is refused at the budget.</summary>
    [Fact]
    public async Task Read_ShapeOfFiftyMillionLetters_IsRefusedWithinTheMemoryTheLimitsPromise()
    {
        byte[] deck = Hostility.WideSlide(2_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(deck), "wide.pptx");

        long peak = heap.Bytes;
        output.WriteLine($"refusing a shape of {deck.Length / 1_048_576} MB took {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.Equal(ReadingFailure.TooLarge, reading.Failure);
        Assert.True(peak < 256L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>
    /// The styles of a workbook are its table of cell formats, which Excel keeps in the tens of thousands. Five million formats are a
    /// hundred megabytes the zip limits let through; read as a tree, the whole table stands in memory before the first cell is read.
    /// </summary>
    [Fact]
    public async Task Read_WorkbookOfFiveMillionCellFormats_IsRefusedWithinTheMemoryTheLimitsPromise()
    {
        byte[] workbook = Hostility.BloatedStyles(5_000_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(workbook), "styles.xlsx");

        long peak = heap.Bytes;
        output.WriteLine($"a workbook of {workbook.Length / 1_048_576} MB of styles was answered '{reading.Failure?.ToString() ?? "read"}' in {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.Equal(ReadingFailure.TooLarge, reading.Failure);
        Assert.True(peak < 512L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>
    /// The list of sheets is the workbook's table of contents, a few lines in a real file. Two million entries are a hundred megabytes the
    /// zip limits let through; read as a tree, the list stands whole, and every sheet in it becomes a part of the document before the
    /// budget of characters is asked.
    /// </summary>
    [Fact]
    public async Task Read_WorkbookListingTwoMillionSheets_IsRefusedWithinTheMemoryTheLimitsPromise()
    {
        byte[] workbook = Hostility.ListOfSheets(2_000_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(workbook), "sheets.xlsx");

        long peak = heap.Bytes;
        output.WriteLine($"a workbook of {workbook.Length / 1_048_576} MB listing sheets was answered '{reading.Failure?.ToString() ?? "read"}' in {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.Equal(ReadingFailure.TooLarge, reading.Failure);
        Assert.True(peak < 512L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>
    /// The list of slides is the presentation's table of contents. Two million entries that name no slide make a document of nothing; read
    /// as a tree, the list stood whole to find that out.
    /// </summary>
    [Fact]
    public async Task Read_PresentationListingTwoMillionSlides_IsAnsweredWithinTheMemoryTheLimitsPromise()
    {
        byte[] deck = Hostility.ListOfSlides(2_000_000);
        GC.Collect();
        using var heap = new PeakHeap();
        var clock = Stopwatch.StartNew();

        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(deck), "slides.pptx");

        long peak = heap.Bytes;
        output.WriteLine($"a presentation of {deck.Length / 1_048_576} MB listing slides was answered '{reading.Failure?.ToString() ?? "read"}' in {clock.Elapsed.TotalSeconds:0.0} s with a heap of {peak / 1_048_576} MB at its peak");
        Assert.NotEqual(ReadingFailure.Damaged, reading.Failure);
        Assert.True(peak < 512L * 1024 * 1024, $"a heap of {peak / 1_048_576} MB at its peak");
    }

    /// <summary>
    /// The separator is guessed from the first line; cut off with <c>Split('\n', 2)</c>, the rest of the file was copied into a second string,
    /// and a file of twelve million characters paid 30 MB to read its first line. A span of the line costs nothing.
    /// </summary>
    [Fact]
    public void Separator_FirstLineOfATwelveMegabyteFile_IsFoundWithoutCopyingTheRest()
    {
        string text = string.Concat(Enumerable.Repeat("a;b;c\n", 2_000_000));
        long before = GC.GetAllocatedBytesForCurrentThread();

        char separator = CsvReader.Separator(text);

        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.Equal(';', separator);
        Assert.True(allocated < 64 * 1024, $"finding the separator allocated {allocated / 1_048_576} MB");
    }

    /// <summary>
    /// The character budget is checked per block after the block is written, and one table of a hundred thousand cells of four hundred
    /// characters was forty million characters in one string before the budget saw it; the characters its cells hold are counted first now.
    /// </summary>
    [Fact]
    public void Of_OneTableTwiceTheCharacterBudget_IsRefusedBeforeItIsWrittenWhole()
    {
        string cell = new('x', 400);
        IReadOnlyList<IReadOnlyList<string>> rows = [.. Enumerable.Range(0, 1_000).Select(_ => (IReadOnlyList<string>)[.. Enumerable.Repeat(cell, 100)])];
        long before = GC.GetAllocatedBytesForCurrentThread();

        Assert.Throws<DocumentTooLargeException>(() => Document.Of(DocumentKind.Csv, [new TableBlock(null, rows)]));

        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(allocated < 2L * ReadingLimits.MostDocumentCharacters, $"refusing one table allocated {allocated / 1_048_576} MB");
    }
}
