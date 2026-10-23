namespace Snail.Toolkit.Documents.Tests.Chaos;

/// <summary>One shared <see cref="DocumentReader.Every"/> read from many threads at once: the readers keep no state, and the answer must not depend on who else is reading.</summary>
public sealed class ParallelChaosTests
{
    [Fact]
    public async Task Read_EveryFormatOnThirtyTwoThreadsAtOnce_ReadsEachAsWhenAlone()
    {
        (string Name, byte[] File)[] files = [("a.docx", Samples.Word()), ("b.xlsx", Samples.ExcelOfEveryCell()), ("c.pptx", Samples.Slides()), ("d.pdf", Samples.Pdf()), ("e.xls", Samples.OldExcel()), ("f.csv", Samples.Csv())];
        var alone = await Task.WhenAll(files.Select(file => ReadAsync(file.Name, file.File)));

        var together = await Task.WhenAll(Enumerable.Range(0, 32).SelectMany(_ => files).Select(file => Task.Run(() => ReadAsync(file.Name, file.File))));

        Assert.All(together.Chunk(files.Length), batch => Assert.Equal(alone, batch));
    }

    /// <summary>Eight hostile workbooks refused at once must each be refused for the reason the budget gives, not for the memory the others took.</summary>
    [Fact]
    public async Task Read_EightSparseWorkbooksAtOnce_AreEachRefusedAsTooLarge()
    {
        byte[] sparse = Hostility.SparseWorkbook();

        var readings = await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => Task.Run(() => DocumentReader.Every.ReadAsync(new MemoryStream(sparse), "sparse.xlsx"))));

        Assert.All(readings, reading => Assert.Equal(ReadingFailure.TooLarge, reading.Failure));
    }

    private static async Task<string> ReadAsync(string name, byte[] file)
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(file), name);

        return reading.Document?.Text ?? $"{reading.Failure}: {reading.Reason}";
    }
}
