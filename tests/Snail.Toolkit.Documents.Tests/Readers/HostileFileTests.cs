using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Snail.Toolkit.Documents.Readers;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Snail.Toolkit.Documents.Tests.Readers;

/// <summary>Files built to hurt the reader — small on disk, enormous or malformed once read — refused with a reason, never followed.</summary>
/// <remarks>
/// The size of a file says nothing about what it becomes. Each of these was found by the audit of Snail.MCP.Documents of
/// 2026-10-16 as a way to take a process down, and the bound that stops it lives in the reader, so every program that
/// reads documents with this library is held by it.
/// </remarks>
public sealed class HostileFileTests
{
    /// <summary>
    /// A cell at column XFD on every row makes every row 16 384 cells wide: a workbook of 11 kB unpacked to 98 million
    /// characters and 4.69 GB before the budget of cells refused it.
    /// </summary>
    [Fact]
    public async Task Read_WorkbookOfSparseRowsAtTheLastColumn_IsRefusedAsTooLarge()
    {
        var rows = Enumerable.Range(1, 200).Select(row => SheetRow((uint)row, Cell(row == 1 ? $"A{row}" : $"XFD{row}", "x")));
        byte[] workbook = Workbook([.. rows]);

        await Assert.ThrowsAsync<DocumentTooLargeException>(() => ReadAsync(new ExcelReader(), workbook));
    }

    /// <summary>Records that each name a key of their own become one table under the union of the keys: n records, an n by n grid.</summary>
    [Fact]
    public async Task Read_JsonRecordsWithAKeyEach_IsRefusedAsTooLarge()
    {
        string records = $"[{string.Join(',', Enumerable.Range(0, 2_000).Select(index => $$"""{"k{{index}}":1}"""))}]";

        await Assert.ThrowsAsync<DocumentTooLargeException>(() => ReadAsync(new JsonReader(), Encoding.UTF8.GetBytes(records)));
    }

    /// <summary>A reference past XFD overflowed the column arithmetic into an array no process can allocate; it is a damaged file.</summary>
    [Fact]
    public async Task Read_CellReferencePastEveryColumn_IsCalledDamaged()
    {
        byte[] workbook = Workbook([SheetRow(1, Cell("ZZZZZZZ1", "x"))]);

        await Assert.ThrowsAsync<InvalidDataException>(() => ReadAsync(new ExcelReader(), workbook));
    }

    /// <summary>
    /// A list level is an integer in the file, and Markdown indents by twice it: a level near int.MaxValue overflowed into
    /// a negative width. Indented no deeper than the eight levels Word numbers, as a heading is no deeper than six.
    /// </summary>
    [Fact]
    public async Task Read_ListItemAtALevelNoWordWrites_IsIndentedAtMostEightLevels()
    {
        byte[] word = Word(new W.Paragraph(
            new W.ParagraphProperties(new W.NumberingProperties(new W.NumberingLevelReference { Val = 1_200_000_000 }, new W.NumberingId { Val = 1 })),
            new W.Run(new W.Text("item"))));

        Document document = await ReadAsync(new WordReader(), word);

        Assert.Equal($"{new string(' ', ReadingLimits.MostListDepth * 2)}1. item", document.Text);
    }

    /// <summary>
    /// The separator of a Markdown table opens with two runs of spaces around an optional bar; backtracking tried every
    /// split of them, and a row followed by 40 000 spaces took 22 s. Five seconds, not one: the bound tells 22 s from
    /// milliseconds, and a tighter one fails under a parallel run for reasons that are not the regex.
    /// </summary>
    [Fact]
    public async Task Read_MarkdownRowFollowedByALongLineOfSpaces_IsReadInSeconds()
    {
        byte[] page = Encoding.UTF8.GetBytes($"| a |\n{new string(' ', 40_000)}x");
        var clock = Stopwatch.StartNew();

        await new MarkdownReader().ReadAsync(new MemoryStream(page));

        Assert.InRange(clock.Elapsed, TimeSpan.Zero, TimeSpan.FromSeconds(5));
    }

    private static async Task<Document> ReadAsync(IFormatReader reader, byte[] file) =>
        Markdown.Of(reader.Format, await reader.ReadAsync(new MemoryStream(file)));

    private static byte[] Workbook(S.Row[] rows)
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();

            var sheet = workbook.AddNewPart<WorksheetPart>();
            sheet.Worksheet = new S.Worksheet(new S.SheetData(rows));

            workbook.Workbook.AppendChild(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1, Name = "Sheet1" }));
        }

        return stream.ToArray();
    }

    private static byte[] Word(params W.Paragraph[] paragraphs)
    {
        using var stream = new MemoryStream();

        using (var word = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            word.AddMainDocumentPart().Document = new W.Document(new W.Body(paragraphs));

        return stream.ToArray();
    }

    private static S.Row SheetRow(uint index, params S.Cell[] cells) =>
        new(cells) { RowIndex = index };

    private static S.Cell Cell(string reference, string text) =>
        new() { CellReference = reference, DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text(text)) };
}
