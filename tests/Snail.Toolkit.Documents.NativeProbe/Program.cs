using DocumentFormat.OpenXml.Packaging;
using Snail.Toolkit.Documents;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;
using W = DocumentFormat.OpenXml.Wordprocessing;

var word = await DocumentReader.Every.ReadAsync(new MemoryStream(Word()), "report.docx");
var pdf = await DocumentReader.Every.ReadAsync(new MemoryStream(Pdf()), "report.pdf");
var unknown = await DocumentReader.Every.ReadAsync(new MemoryStream([1, 2, 3]), "letter.rtf");

Console.WriteLine(word.Document?.Text ?? word.Reason);
Console.WriteLine(pdf.Document?.Text ?? pdf.Reason);

bool isRead = word.Document?.Text.Contains("Paid in full", StringComparison.Ordinal) == true
    && pdf.Document?.Text.Contains("Quarterly", StringComparison.Ordinal) == true
    && unknown.Failure == ReadingFailure.UnknownFormat;

return isRead ? 0 : 1;

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
