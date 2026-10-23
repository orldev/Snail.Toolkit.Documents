using BenchmarkDotNet.Attributes;
using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents.Benchmarks;

/// <summary>What documents_read does with a file once it is in memory: the format read into blocks, the blocks into Markdown.</summary>
[MemoryDiagnoser]
public class ReadingBenchmarks
{
    private readonly WordReader _word = new();

    private readonly ExcelReader _excel = new();

    private readonly HtmlReader _html = new();

    private readonly MarkdownReader _markdown = new();

    private readonly CsvReader _csv = new();

    private readonly JsonReader _json = new();

    private readonly PdfReader _pdf = new();

    private readonly PowerPointReader _powerPoint = new();

    private byte[] _report = [];

    private byte[] _ledger = [];

    private byte[] _page = [];

    private byte[] _notes = [];

    private byte[] _export = [];

    private byte[] _records = [];

    private byte[] _pages = [];

    private byte[] _deck = [];

    [GlobalSetup]
    public void Setup()
    {
        _report = Files.Report();
        _ledger = Files.Ledger();
        _page = Files.Page();
        _notes = Files.Notes();
        _export = Files.Export();
        _records = Files.Records();
        _pages = Files.Pages();
        _deck = Files.Deck();
    }

    /// <summary>A Word report of twenty sections read as Markdown.</summary>
    [Benchmark]
    public Task<int> Word() => Read(_word, _report);

    /// <summary>A workbook of three sheets of four hundred rows read as Markdown.</summary>
    [Benchmark]
    public Task<int> Excel() => Read(_excel, _ledger);

    /// <summary>A web page of twenty sections in nested containers read as Markdown.</summary>
    [Benchmark]
    public Task<int> Html() => Read(_html, _page);

    /// <summary>Notes of twenty sections in Markdown read as Markdown.</summary>
    [Benchmark]
    public Task<int> Notes() => Read(_markdown, _notes);

    /// <summary>A CSV export of two thousand rows read as one table.</summary>
    [Benchmark]
    public Task<int> Csv() => Read(_csv, _export);

    /// <summary>A JSON export of a thousand records read as one table.</summary>
    [Benchmark]
    public Task<int> Json() => Read(_json, _records);

    /// <summary>A PDF of twenty pages of text read as Markdown.</summary>
    [Benchmark]
    public Task<int> Pdf() => Read(_pdf, _pages);

    /// <summary>A deck of twenty slides with points, a table and notes read as Markdown.</summary>
    [Benchmark]
    public Task<int> PowerPoint() => Read(_powerPoint, _deck);

    private static async Task<int> Read(IFormatReader reader, byte[] file)
    {
        using var content = new MemoryStream(file);
        var blocks = await reader.ReadAsync(content);

        return Document.Of(reader.Kind, blocks).Text.Length;
    }
}
