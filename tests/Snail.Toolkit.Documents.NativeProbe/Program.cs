using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Snail.Toolkit.Documents;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

(string Name, byte[] File, string Expected)[] files =
[
    ("report.docx", Word(), "Paid in full"),
    ("report.pdf", Pdf(), "Quarterly"),
    ("ledger.xlsx", Excel(), "| 1 | Amount |"),
    ("invoices.xls", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "invoices.xls")), "INV-7"),
    ("deck.pptx", Slides(), "Plan"),
    ("page.html", Encoding.UTF8.GetBytes("<html><body><h2>Prices</h2><p>Bread 50</p><ul><li>one</li></ul><table><tr><td>a</td></tr></table></body></html>"), "Bread 50"),
    ("rows.csv", Encoding.UTF8.GetBytes("City;Sum\nMoscow;\"1;5\"\n"), "| Moscow | 1;5 |"),
    ("rows.tsv", Encoding.UTF8.GetBytes("City\tSum\nMoscow\t7\n"), "| Moscow | 7 |"),
    ("records.json", Encoding.UTF8.GetBytes("""[{"name":"a","n":1}]"""), "| a | 1 |"),
    ("notes.md", Encoding.UTF8.GetBytes("# Title\n\n- item\n\n| a | b |\n|---|---|\n| 1 | 2 |\n"), "# Title"),
    ("plain.txt", Encoding.UTF8.GetBytes("First\n\nSecond"), "Second")
];

bool isEveryRead = true;

foreach (var (name, file, expected) in files)
{
    var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(file), name);
    bool isRead = reading.Document?.Text.Contains(expected, StringComparison.Ordinal) == true;
    isEveryRead &= isRead;
    Console.WriteLine($"{(isRead ? "ok  " : "FAIL")} {name}: {(reading.IsRead ? $"{reading.Document!.Text.Length} chars" : $"{reading.Failure}: {reading.Reason}")}");
}

var unknown = await DocumentReader.Every.ReadAsync(new MemoryStream([1, 2, 3]), "letter.rtf");
var scans = ScannedPages.Pictures(new MemoryStream(Pdf()), [1]);
Console.WriteLine($"unknown: {unknown.Failure}; scans: {scans.Count}");

return isEveryRead && unknown.Failure == ReadingFailure.UnknownFormat ? 0 : 1;

static byte[] Word()
{
    using var stream = new MemoryStream();

    using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        document.AddMainDocumentPart().Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text("Paid in full")))));

    return stream.ToArray();
}

static byte[] Pdf()
{
    var builder = new PdfDocumentBuilder();
    var font = builder.AddStandard14Font(Standard14Font.Helvetica);
    builder.AddPage(PageSize.A4).AddText("Quarterly report", 12, new PdfPoint(72, 720), font);

    return builder.Build();
}

static byte[] Excel()
{
    using var stream = new MemoryStream();

    using (var book = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
    {
        var workbook = book.AddWorkbookPart();
        workbook.Workbook = new S.Workbook();
        var sheet = workbook.AddNewPart<WorksheetPart>();
        var header = new S.Cell { CellReference = "A1", DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text("Amount")) };
        var amount = new S.Cell { CellReference = "B1", CellValue = new S.CellValue("12.5") };
        var row = new S.Row(header, amount) { RowIndex = 1 };
        sheet.Worksheet = new S.Worksheet(new S.SheetData(row));
        workbook.Workbook.AppendChild(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1, Name = "Sheet1" }));
    }

    return stream.ToArray();
}

static byte[] Slides()
{
    using var stream = new MemoryStream();

    using (var deck = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
    {
        var presentation = deck.AddPresentationPart();
        var slide = presentation.AddNewPart<SlidePart>();
        var title = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2, Name = "Title" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
            new P.ShapeProperties(),
            new P.TextBody(new A.BodyProperties(), new A.Paragraph(new A.Run(new A.Text("Plan")))));
        slide.Slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree(title)));
        presentation.Presentation = new P.Presentation(new P.SlideIdList(new P.SlideId { Id = 256, RelationshipId = presentation.GetIdOfPart(slide) }));
    }

    return stream.ToArray();
}
