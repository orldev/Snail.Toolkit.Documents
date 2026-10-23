using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Snail.Toolkit.Documents.Benchmarks;

/// <summary>Documents of the size a task brings, built in memory the way their own programs write them.</summary>
public static class Files
{
    /// <summary>A report of twenty sections, each a heading, eight paragraphs, a bulleted list and a table of twelve rows.</summary>
    public static byte[] Report()
    {
        using var stream = new MemoryStream();

        using (var word = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = word.AddMainDocumentPart();

            main.AddNewPart<StyleDefinitionsPart>().Styles = new W.Styles(
                new W.Style(new W.StyleName { Val = "heading 1" }) { Type = W.StyleValues.Paragraph, StyleId = "Heading1" });

            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new W.Numbering(
                new W.AbstractNum(new W.Level(new W.NumberingFormat { Val = W.NumberFormatValues.Bullet }) { LevelIndex = 0 }) { AbstractNumberId = 0 },
                new W.NumberingInstance(new W.AbstractNumId { Val = 0 }) { NumberID = 1 });

            var body = new W.Body();

            for (var section = 1; section <= 20; section++)
            {
                body.Append(Paragraph($"Section {section}: what the week brought", new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Heading1" })));

                for (var paragraph = 1; paragraph <= 8; paragraph++)
                    body.Append(Paragraph($"Paragraph {paragraph} of section {section}: the work went as planned, the numbers below say how far, and the risks are named at the end."));

                for (var point = 1; point <= 4; point++)
                    body.Append(Paragraph($"point {point}", new W.ParagraphProperties(new W.NumberingProperties(new W.NumberingLevelReference { Val = 0 }, new W.NumberingId { Val = 1 }))));

                var table = new W.Table(Row("Month", "Planned", "Done", "Note"));

                for (var row = 1; row <= 12; row++)
                    table.Append(Row($"Month {row}", (row * 100).ToString(CultureInfo.InvariantCulture), (row * 97).ToString(CultureInfo.InvariantCulture), "on track"));

                body.Append(table);
            }

            main.Document = new W.Document(body);
        }

        return stream.ToArray();
    }

    /// <summary>A workbook of three sheets of four hundred rows and six columns: text, numbers and dates.</summary>
    public static byte[] Ledger()
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();

            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new S.SharedStringTable(
                new S.SharedStringItem(new S.Text("Date")),
                new S.SharedStringItem(new S.Text("Account")),
                new S.SharedStringItem(new S.Text("Amount")));

            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new S.Stylesheet(
                new S.CellFormats(new S.CellFormat { NumberFormatId = 0 }, new S.CellFormat { NumberFormatId = 14, ApplyNumberFormat = true }));

            var sheets = new S.Sheets();

            for (uint sheet = 1; sheet <= 3; sheet++)
            {
                var data = new S.SheetData(new S.Row(Shared("A1", 0), Shared("B1", 1), Shared("C1", 2)) { RowIndex = 1 });

                for (uint row = 2; row <= 401; row++)
                {
                    var line = new S.Row(
                        new S.Cell { CellReference = $"A{row}", StyleIndex = 1, CellValue = new S.CellValue((46000 + row).ToString(CultureInfo.InvariantCulture)) },
                        new S.Cell { CellReference = $"B{row}", DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text($"account {row % 17}")) },
                        Number($"C{row}", (row * 1.1).ToString(CultureInfo.InvariantCulture)),
                        Number($"D{row}", (row * 3).ToString(CultureInfo.InvariantCulture)),
                        Number($"E{row}", "0.30000000000000004"),
                        Number($"F{row}", (row % 5).ToString(CultureInfo.InvariantCulture)))
                    {
                        RowIndex = row
                    };

                    data.Append(line);
                }

                var part = workbook.AddNewPart<WorksheetPart>();
                part.Worksheet = new S.Worksheet(data);
                sheets.Append(new S.Sheet { Id = workbook.GetIdOfPart(part), SheetId = sheet, Name = $"Sheet {sheet}" });
            }

            workbook.Workbook.AppendChild(sheets);
        }

        return stream.ToArray();
    }

    /// <summary>A deck of twenty slides: a title, six points on two levels, a table of six rows and the speaker's notes each.</summary>
    public static byte[] Deck()
    {
        using var stream = new MemoryStream();

        using (var deck = PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
        {
            var presentation = deck.AddPresentationPart();
            var ids = new P.SlideIdList();

            for (uint number = 1; number <= 20; number++)
            {
                var slide = presentation.AddNewPart<SlidePart>();
                var table = new A.Table(CellRow("Month", "Planned", "Done"));

                for (int row = 1; row <= 6; row++)
                    table.Append(CellRow($"Month {row}", (row * 100).ToString(CultureInfo.InvariantCulture), (row * 97).ToString(CultureInfo.InvariantCulture)));

                slide.Slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree(
                    Placeholder(P.PlaceholderValues.Title, Text($"Slide {number}: what the week brought")),
                    Placeholder(P.PlaceholderValues.Body, Text("first point", 0), Text("a detail", 1), Text("second point", 0), Text("another detail", 1), Text("third point", 0), Text("last detail", 1)),
                    new P.GraphicFrame(new A.Graphic(new A.GraphicData(table))))));

                slide.AddNewPart<NotesSlidePart>().NotesSlide = new P.NotesSlide(new P.CommonSlideData(new P.ShapeTree(
                    Placeholder(P.PlaceholderValues.Body, Text($"Say what slide {number} leaves out, and name the risks.")))));

                ids.Append(new P.SlideId { Id = 255 + number, RelationshipId = presentation.GetIdOfPart(slide) });
            }

            presentation.Presentation = new P.Presentation(ids);
        }

        return stream.ToArray();
    }

    /// <summary>A saved web page of twenty sections inside the nested containers a site builder writes: headings, paragraphs, a list and a table each.</summary>
    public static byte[] Page()
    {
        var page = new StringBuilder("<html><head><title>Weekly report</title><style>p{}</style></head><body><nav><a href=\"/\">Home</a></nav>");

        for (int section = 1; section <= 20; section++)
        {
            page.Append("<div class=\"row\"><div class=\"col\"><div class=\"card\"><div class=\"body\">");
            page.Append($"<h2>Section {section}: what the week brought</h2>");

            for (int paragraph = 1; paragraph <= 8; paragraph++)
                page.Append($"<p>Paragraph {paragraph} of section {section}: the work went <b>as planned</b>, the numbers below say how far.</p>");

            page.Append("<ul><li>first point<ul><li>a detail</li></ul></li><li>second point</li></ul>");
            page.Append("<table><tr><th>Month</th><th>Planned</th><th>Done</th></tr>");

            for (int row = 1; row <= 12; row++)
                page.Append($"<tr><td>Month {row}</td><td>{row * 100}</td><td>{row * 97}</td></tr>");

            page.Append("</table></div></div></div></div>");
        }

        return Encoding.UTF8.GetBytes(page.Append("<script>track()</script></body></html>").ToString());
    }

    /// <summary>Notes in Markdown of twenty sections: a heading, paragraphs, a list, a table and a fenced block each.</summary>
    public static byte[] Notes()
    {
        var notes = new StringBuilder();

        for (int section = 1; section <= 20; section++)
        {
            notes.Append($"## Section {section}\n\n");

            for (int paragraph = 1; paragraph <= 8; paragraph++)
                notes.Append($"Paragraph {paragraph} of section {section}: the work went as planned, and the numbers say how far.\n\n");

            notes.Append("- first point\n  - a detail\n- second point\n\n| Month | Planned | Done |\n|---|---|---|\n");

            for (int row = 1; row <= 12; row++)
                notes.Append($"| Month {row} | {row * 100} | {row * 97} |\n");

            notes.Append("\n```cs\nvar total = rows.Sum();\n```\n\n");
        }

        return Encoding.UTF8.GetBytes(notes.ToString());
    }

    /// <summary>An export of two thousand rows and five columns, semicolons and quoted cells, as a Russian Excel saves one.</summary>
    public static byte[] Export()
    {
        var export = new StringBuilder("Date;Account;Amount;Note;Owner\n");

        for (int row = 1; row <= 2_000; row++)
            export.Append($"2026-10-{row % 28 + 1:00};account {row % 17};{row * 1.1:0.00};\"note; with a separator\";owner {row % 5}\n");

        return Encoding.UTF8.GetBytes(export.ToString());
    }

    /// <summary>A JSON export of a thousand flat records of six fields.</summary>
    public static byte[] Records()
    {
        var records = Enumerable.Range(1, 1_000).Select(row => $$"""{"date":"2026-10-{{row % 28 + 1:00}}","account":"account {{row % 17}}","amount":{{row * 3}},"rate":0.25,"closed":{{(row % 2 == 0 ? "true" : "false")}},"note":null}""");

        return Encoding.UTF8.GetBytes($"[{string.Join(',', records)}]");
    }

    /// <summary>A PDF of twenty pages of text, forty lines each.</summary>
    public static byte[] Pages()
    {
        var builder = new PdfDocumentBuilder();
        var font = builder.AddStandard14Font(Standard14Font.Helvetica);

        for (int number = 1; number <= 20; number++)
        {
            var page = builder.AddPage(UglyToad.PdfPig.Content.PageSize.A4);

            for (int line = 0; line < 40; line++)
                page.AddText($"Page {number}, line {line}: the work went as planned.", 10, new UglyToad.PdfPig.Core.PdfPoint(50, 800 - (line * 18)), font);
        }

        return builder.Build();
    }

    /// <summary>A workbook of two thousand rows that each hold one cell at XFD, the last column: eleven kilobytes on disk.</summary>
    public static byte[] Sparse()
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();
            var data = new S.SheetData();

            for (uint row = 1; row <= 2_000; row++)
                data.Append(new S.Row(Number(row == 1 ? $"A{row}" : $"XFD{row}", "1")) { RowIndex = row });

            var part = workbook.AddNewPart<WorksheetPart>();
            part.Worksheet = new S.Worksheet(data);
            workbook.Workbook.AppendChild(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(part), SheetId = 1, Name = "Sheet 1" }));
        }

        return stream.ToArray();
    }

    /// <summary>Two thousand JSON records that each name a key of their own: twenty-three kilobytes on disk.</summary>
    public static byte[] Keys() =>
        Encoding.UTF8.GetBytes($"[{string.Join(',', Enumerable.Range(0, 2_000).Select(index => $$"""{"k{{index}}":1}"""))}]");

    private static P.Shape Placeholder(P.PlaceholderValues type, params A.Paragraph[] paragraphs) =>
        new(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2, Name = type.ToString() },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = type })),
            new P.ShapeProperties(),
            new P.TextBody([new A.BodyProperties(), .. paragraphs]));

    private static A.Paragraph Text(string text, int level = 0) =>
        new(new A.ParagraphProperties { Level = level }, new A.Run(new A.Text(text)));

    private static A.TableRow CellRow(params string[] cells) =>
        new([.. cells.Select(cell => (OpenXmlElement)new A.TableCell(new A.TextBody(new A.BodyProperties(), Text(cell))))]);

    private static W.Paragraph Paragraph(string text, W.ParagraphProperties? properties = null) =>
        properties is null
            ? new W.Paragraph(new W.Run(new W.Text(text)))
            : new W.Paragraph(properties, new W.Run(new W.Text(text)));

    private static W.TableRow Row(params string[] cells) =>
        new([.. cells.Select(cell => new W.TableCell(Paragraph(cell)))]);

    private static S.Cell Shared(string reference, int index) =>
        new() { CellReference = reference, DataType = S.CellValues.SharedString, CellValue = new S.CellValue(index.ToString(CultureInfo.InvariantCulture)) };

    private static S.Cell Number(string reference, string value) =>
        new() { CellReference = reference, CellValue = new S.CellValue(value) };
}
