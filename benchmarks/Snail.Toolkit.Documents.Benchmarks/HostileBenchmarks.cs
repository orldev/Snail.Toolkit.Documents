using BenchmarkDotNet.Attributes;
using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents.Benchmarks;

/// <summary>What a file built to amplify costs before it is refused: the budget of a document is a ceiling on memory only if it holds here.</summary>
/// <remarks>Before the budget, the sparse workbook came to 98 408 325 characters and 4.69 GB allocated; a grown allocation here is the budget leaking.</remarks>
[MemoryDiagnoser]
public class HostileBenchmarks
{
    private byte[] _sparse = [];

    private byte[] _keys = [];

    [GlobalSetup]
    public void Setup()
    {
        _sparse = Files.Sparse();
        _keys = Files.Keys();
    }

    /// <summary>A workbook of two thousand rows each reaching the last column, read until the budget refuses it.</summary>
    [Benchmark]
    public Task<bool> IsSparseSheetRefused() => IsRefusedAsync(new ExcelReader(), _sparse);

    /// <summary>Two thousand JSON records under a key each, read until the budget refuses them.</summary>
    [Benchmark]
    public Task<bool> IsKeyPerRecordRefused() => IsRefusedAsync(new JsonReader(), _keys);

    private static async Task<bool> IsRefusedAsync(IFormatReader reader, byte[] file)
    {
        try
        {
            using var content = new MemoryStream(file);
            Markdown.Of(reader.Format, await reader.ReadAsync(content));

            return false;
        }
        catch (DocumentTooLargeException)
        {
            return true;
        }
    }
}
