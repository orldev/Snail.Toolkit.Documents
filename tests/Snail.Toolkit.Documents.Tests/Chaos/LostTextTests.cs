using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Snail.Toolkit.Documents.Readers;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace Snail.Toolkit.Documents.Tests.Chaos;

/// <summary>Text a real document holds that the reader lost, doubled or split wrongly: a model answers from the text it was given.</summary>
public sealed class LostTextTests
{
    /// <summary>
    /// Word writes a text box inside <c>mc:AlternateContent</c>: once as a shape for programs that know shapes (<c>mc:Choice</c>) and
    /// once as the older VML drawing (<c>mc:Fallback</c>). A walk over <c>Descendants()</c> enters both, and every text box in a
    /// document reads twice over; with the choices resolved as Word resolves them, it reads once.
    /// </summary>
    [Fact]
    public async Task Word_TextBoxWrittenWithItsFallback_KeepsItsTextOnce()
    {
        var shape = new Wps.WordprocessingShape(new Wps.TextBoxInfo2(new W.TextBoxContent(Paragraph("Boxed once"))));
        var drawing = new W.Drawing(new Wp.Inline(new A.Graphic(new A.GraphicData(shape) { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" })));
        var fallback = new W.Picture(new V.Shape(new V.TextBox(new W.TextBoxContent(Paragraph("Boxed once")))));
        var alternatives = new AlternateContent(new AlternateContentChoice(drawing) { Requires = "wps" }, new AlternateContentFallback(fallback));
        alternatives.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        var boxed = new W.Paragraph(new W.Run(alternatives));

        var document = await ReadAsync(Word(boxed));

        Assert.Equal(1, Count(document.Text, "Boxed once"));
    }

    /// <summary>Read as its direct paragraphs only, a cell holding a content control or a nested table came out empty where it mattered; a cell is walked like the body.</summary>
    [Fact]
    public async Task Word_TableCellHoldingAContentControlAndANestedTable_KeepsTheirText()
    {
        var table = new W.Table(new W.TableRow(
            new W.TableCell(new W.SdtBlock(new W.SdtProperties(), new W.SdtContentBlock(Paragraph("Inside the control")))),
            new W.TableCell(Paragraph("Outer"), new W.Table(new W.TableRow(new W.TableCell(Paragraph("Nested cell")))))));

        var document = await ReadAsync(Word(table));

        Assert.Contains("Inside the control", document.Text, StringComparison.Ordinal);
        Assert.Contains("Nested cell", document.Text, StringComparison.Ordinal);
    }

    /// <summary>
    /// A reader that walks a part's events reads an element to its end; an empty element written as <c>&lt;w:p/&gt;</c> must end where it
    /// starts, or the walk would swallow the paragraph after it.
    /// </summary>
    [Fact]
    public async Task Word_EmptyParagraphsAndAnEmptyTable_LoseNothingAroundThem()
    {
        var document = await ReadAsync(Word(Paragraph("Before"), new W.Paragraph(), new W.Table(), new W.Paragraph(new W.ParagraphProperties()), Paragraph("After")));

        Assert.Equal("Before\n\nAfter", document.Text);
    }

    /// <summary>An empty cell written as <c>&lt;c r="B1"/&gt;</c> ends where it starts, so the cell after it keeps its column.</summary>
    [Fact]
    public async Task Sheet_EmptyCellBetweenFullOnes_KeepsItsNeighboursInPlace()
    {
        var document = Document.Of(DocumentKind.Excel, await new ExcelReader().ReadAsync(new MemoryStream(SheetWithEmptyCells())));

        Assert.Equal([["a", string.Empty, "c"]], document.Tables[0].Rows.Select(row => row.ToArray()));
    }

    /// <summary>An ASP.NET WebForms page wraps its whole body in a form; skipped as a control, such a page came out empty.</summary>
    [Fact]
    public async Task Html_PageWhoseWholeContentIsInsideAForm_KeepsTheContent()
    {
        byte[] page = Encoding.UTF8.GetBytes("<html><body><form method=\"post\" action=\"./Default.aspx\"><h1>Prices</h1><p>Bread 50</p></form></body></html>");

        var document = Document.Of(DocumentKind.Html, await new HtmlReader().ReadAsync(new MemoryStream(page)));

        Assert.Contains("Bread 50", document.Text, StringComparison.Ordinal);
    }

    /// <summary>A file called .tsv says its separator in its name; guessed from the first line instead, a header with more commas than tabs split by its commas.</summary>
    [Fact]
    public async Task Csv_TsvWhoseFirstLineHoldsMoreCommasThanTabs_IsSplitByTabs()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream("Name, first, last\tAge\nAnn, B, C\t7\n"u8.ToArray()), "people.tsv");

        Assert.Equal(2, Assert.Single(reading.Document!.Tables).Columns);
    }

    private static byte[] SheetWithEmptyCells()
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();
            var sheet = workbook.AddNewPart<WorksheetPart>();
            var first = new S.Cell { CellReference = "A1", DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text("a")) };
            var third = new S.Cell { CellReference = "C1", DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text("c")) };
            var row = new S.Row(first, new S.Cell { CellReference = "B1" }, third) { RowIndex = 1 };
            sheet.Worksheet = new S.Worksheet(new S.SheetData(row));
            workbook.Workbook.AppendChild(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1, Name = "Sheet1" }));
        }

        return stream.ToArray();
    }

    private static async Task<Document> ReadAsync(byte[] word) =>
        Document.Of(DocumentKind.Word, await new WordReader().ReadAsync(new MemoryStream(word)));

    private static int Count(string text, string piece) =>
        (text.Length - text.Replace(piece, string.Empty, StringComparison.Ordinal).Length) / piece.Length;

    private static W.Paragraph Paragraph(string text) => new(new W.Run(new W.Text(text)));

    private static byte[] Word(params OpenXmlElement[] body)
    {
        using var stream = new MemoryStream();

        using (var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            word.AddMainDocumentPart().Document = new W.Document(new W.Body(body));

        return stream.ToArray();
    }
}
