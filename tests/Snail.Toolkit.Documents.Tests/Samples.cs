using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Snail.Toolkit.Documents.Tests;

/// <summary>Real files of every format, built the way their own programs write them, small enough to read at a glance.</summary>
internal static class Samples
{
    /// <summary>A Word report whose heading style has the id a Russian Word gives «Heading 1»: <c>1</c>, and whose numbering names a list it does not define.</summary>
    public static byte[] Word()
    {
        using var stream = new MemoryStream();

        using (var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = word.AddMainDocumentPart();

            main.AddNewPart<StyleDefinitionsPart>().Styles = new W.Styles(
                new W.Style(new W.StyleName { Val = "heading 1" }) { Type = W.StyleValues.Paragraph, StyleId = "1" });

            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new W.Numbering(
                new W.AbstractNum(new W.Level(new W.NumberingFormat { Val = W.NumberFormatValues.Bullet }) { LevelIndex = 0 }) { AbstractNumberId = 0 },
                new W.NumberingInstance(new W.AbstractNumId { Val = 0 }) { NumberID = 1 },
                new W.NumberingInstance(new W.AbstractNumId { Val = 9 }) { NumberID = 2 });

            main.Document = new W.Document(new W.Body(
                Paragraph("Отчёт за неделю", new W.ParagraphProperties(new W.ParagraphStyleId { Val = "1" })),
                Paragraph("Первый абзац отчёта."),
                Paragraph("пункт списка", new W.ParagraphProperties(new W.NumberingProperties(new W.NumberingLevelReference { Val = 0 }, new W.NumberingId { Val = 1 }))),
                new W.Table(
                    Row("Месяц", "Сумма"),
                    Row("Январь", "100"))));
        }

        return stream.ToArray();
    }

    /// <summary>
    /// A Word document of the shapes a report takes besides styled headings: a title, a paragraph that is a heading by
    /// its own outline level, a style that is one by its level, a content control, tabs and line breaks, an empty
    /// paragraph — and no numbering part at all.
    /// </summary>
    public static byte[] WordOfEveryShape()
    {
        using var stream = new MemoryStream();

        using (var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = word.AddMainDocumentPart();

            main.AddNewPart<StyleDefinitionsPart>().Styles = new W.Styles(
                new W.Style(new W.StyleName { Val = "Title" }) { Type = W.StyleValues.Paragraph, StyleId = "Title" },
                new W.Style(new W.StyleName { Val = "Подраздел" }, new W.StyleParagraphProperties(new W.OutlineLevel { Val = 2 })) { Type = W.StyleValues.Paragraph, StyleId = "Sub" },
                new W.Style(new W.StyleName { Val = "heading 4" }) { Type = W.StyleValues.Paragraph });

            main.Document = new W.Document(new W.Body(
                Paragraph("Заголовок документа", new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Title" })),
                Paragraph("Раздел", new W.ParagraphProperties(new W.OutlineLevel { Val = 1 })),
                Paragraph("Подраздел", new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Sub" })),
                new W.SdtBlock(new W.SdtProperties(), new W.SdtContentBlock(Paragraph("Текст в элементе управления"))),
                new W.Paragraph(new W.Run(new W.Text("Колонка"), new W.TabChar(), new W.Text("значение"), new W.Break(), new W.Text("вторая строка"))),
                new W.Paragraph(new W.Run(new W.Text("   "))),
                Paragraph("пункт без определения списка", new W.ParagraphProperties(new W.NumberingProperties(new W.NumberingLevelReference { Val = 0 }, new W.NumberingId { Val = 1 })))));
        }

        return stream.ToArray();
    }

    /// <summary>A workbook of two sheets: one with a header, a date, a sum and a number stored with its binary noise; one empty.</summary>
    public static byte[] Excel()
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();

            workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new S.SharedStringTable(
                new S.SharedStringItem(new S.Text("Дата")),
                new S.SharedStringItem(new S.Text("Сумма")));

            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new S.Stylesheet(
                new S.CellFormats(new S.CellFormat { NumberFormatId = 0 }, new S.CellFormat { NumberFormatId = 14, ApplyNumberFormat = true }));

            var revenue = workbook.AddNewPart<WorksheetPart>();
            revenue.Worksheet = new S.Worksheet(new S.SheetData(
                new S.Row(Shared("A1", 0), Shared("B1", 1)) { RowIndex = 1 },
                new S.Row(new S.Cell { CellReference = "A2", StyleIndex = 1, CellValue = new S.CellValue("46305") }, Number("B2", "0.30000000000000004")) { RowIndex = 2 },
                new S.Row(Number("B4", "12")) { RowIndex = 4 }));

            var empty = workbook.AddNewPart<WorksheetPart>();
            empty.Worksheet = new S.Worksheet(new S.SheetData());

            workbook.Workbook.AppendChild(new S.Sheets(
                new S.Sheet { Id = workbook.GetIdOfPart(revenue), SheetId = 1, Name = "Выручка" },
                new S.Sheet { Id = workbook.GetIdOfPart(empty), SheetId = 2, Name = "Пусто" }));
        }

        return stream.ToArray();
    }

    /// <summary>
    /// A sheet of every kind of cell: an inline string, two booleans, an error, text where a number should be, a date by a
    /// custom format, a number whose format quotes date letters, a serial too large to be a date, an empty row and a cell
    /// that names no column.
    /// </summary>
    public static byte[] ExcelOfEveryCell() =>
        Workbook(
            new S.Stylesheet(
                new S.NumberingFormats(
                    new S.NumberingFormat { NumberFormatId = 164, FormatCode = "dd.mm.yyyy" },
                    new S.NumberingFormat { NumberFormatId = 165, FormatCode = "0.0\" dm\"" }),
                new S.CellFormats(new S.CellFormat { NumberFormatId = 0 }, new S.CellFormat { NumberFormatId = 164 }, new S.CellFormat { NumberFormatId = 165 })),
            SheetRow(1,
                new S.Cell { CellReference = "A1", DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text("Строка")) },
                new S.Cell { CellReference = "B1", DataType = S.CellValues.Boolean, CellValue = new S.CellValue("1") },
                new S.Cell { CellReference = "C1", DataType = S.CellValues.Boolean, CellValue = new S.CellValue("0") },
                new S.Cell { CellReference = "D1", DataType = S.CellValues.Error, CellValue = new S.CellValue("#DIV/0!") }),
            SheetRow(2,
                Number("A2", "n/a"),
                new S.Cell { CellReference = "B2", StyleIndex = 1, CellValue = new S.CellValue("46305") },
                new S.Cell { CellReference = "C2", StyleIndex = 2, CellValue = new S.CellValue("12.5") },
                new S.Cell { CellReference = "D2", StyleIndex = 1, CellValue = new S.CellValue("3000000") }),
            SheetRow(3, Number("A3", string.Empty)),
            SheetRow(4, new S.Cell { CellValue = new S.CellValue("без адреса") }));

    /// <summary>A workbook saved with no styles part: a cell points at a format the file does not have.</summary>
    public static byte[] ExcelWithoutStyles() =>
        Workbook(null, SheetRow(1, new S.Cell { CellReference = "A1", StyleIndex = 1, CellValue = new S.CellValue("46305") }));

    /// <summary>
    /// An Excel 97–2003 workbook, as accounting systems still export: a sheet of invoices whose cells start at B2, with a
    /// date, a fraction and an empty row 4, and an empty second sheet. No .NET library writes the format, so the file was
    /// written once by xlwt 1.3.0 and is kept beside the tests.
    /// </summary>
    public static byte[] OldExcel() =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "invoices.xls"));

    /// <summary>A presentation of one slide: a title, two points one level apart, a table and the speaker's notes.</summary>
    public static byte[] Slides()
    {
        using var stream = new MemoryStream();

        using (var deck = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentation = deck.AddPresentationPart();
            var slide = presentation.AddNewPart<SlidePart>();

            slide.Slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree(
                Placeholder(P.PlaceholderValues.Title, Text("План работ")),
                Placeholder(P.PlaceholderValues.Body, Text("Первое", 0), Text("Второе", 1)),
                new P.GraphicFrame(new A.Graphic(new A.GraphicData(new A.Table(
                    CellRow("Этап", "Срок"),
                    CellRow("Сбор", "май"))))))));

            slide.AddNewPart<NotesSlidePart>().NotesSlide = new P.NotesSlide(new P.CommonSlideData(new P.ShapeTree(
                Placeholder(P.PlaceholderValues.Body, Text("Сказать про сроки")))));

            presentation.Presentation = new P.Presentation(new P.SlideIdList(new P.SlideId { Id = 256, RelationshipId = presentation.GetIdOfPart(slide) }));
        }

        return stream.ToArray();
    }

    /// <summary>
    /// A presentation whose first slide is listed without a link to its part, and whose second holds a text box inside a
    /// group, with a line break and an empty line in it.
    /// </summary>
    public static byte[] SlidesOfEveryShape()
    {
        using var stream = new MemoryStream();

        using (var deck = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentation = deck.AddPresentationPart();
            var slide = presentation.AddNewPart<SlidePart>();

            var box = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 3, Name = "Text box" },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.Paragraph(new A.Run(new A.Text("Первая")), new A.Break(), new A.Run(new A.Text("вторая"))),
                    new A.Paragraph()));

            slide.Slide = new P.Slide(new P.CommonSlideData(new P.ShapeTree(
                Placeholder(P.PlaceholderValues.Title, Text("Схема")),
                new P.GroupShape(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 4, Name = "Group" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties(),
                    box))));

            presentation.Presentation = new P.Presentation(new P.SlideIdList(
                new P.SlideId { Id = 256 },
                new P.SlideId { Id = 257, RelationshipId = presentation.GetIdOfPart(slide) }));
        }

        return stream.ToArray();
    }

    /// <summary>A PDF of two pages: text on the first, a picture alone on the second — a scan.</summary>
    public static byte[] Pdf()
    {
        var builder = new PdfDocumentBuilder();
        var font = builder.AddStandard14Font(Standard14Font.Helvetica);

        builder.AddPage(UglyToad.PdfPig.Content.PageSize.A4).AddText("Quarterly report", 12, new UglyToad.PdfPig.Core.PdfPoint(72, 720), font);
        builder.AddPage(UglyToad.PdfPig.Content.PageSize.A4).AddPng(Png, new UglyToad.PdfPig.Core.PdfRectangle(72, 72, 300, 300));

        return builder.Build();
    }

    /// <summary>A PDF of one page holding one grey picture of <paramref name="width"/> by <paramref name="height"/> pixels, written as a PNG by hand.</summary>
    /// <remarks>Rows of zeros deflate to almost nothing, so the file is small and the picture is as large as its dimensions say.</remarks>
    public static byte[] PdfOfOneHugePicture(int width, int height)
    {
        var builder = new PdfDocumentBuilder();
        builder.AddPage(UglyToad.PdfPig.Content.PageSize.A4).AddPng(GreyPng(width, height), new UglyToad.PdfPig.Core.PdfRectangle(20, 20, 575, 820));

        return builder.Build();
    }

    /// <summary>A PDF locked by a password nobody here knows: its trailer names a standard security handler that rejects the empty one.</summary>
    public static byte[] LockedPdf()
    {
        string pdf = Encoding.Latin1.GetString(Pdf());
        string key = new('0', 64);

        string locked = pdf.Replace("trailer\n<<", $"trailer\n<< /Encrypt << /Filter /Standard /V 1 /R 2 /O <{key}> /U <{key}> /P -4 >> ", StringComparison.Ordinal);

        return Encoding.Latin1.GetBytes(locked);
    }

    /// <summary>A CSV the way a Russian Excel saves one: Windows-1251, semicolons, a quoted cell holding the separator.</summary>
    public static byte[] Csv()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        return Encoding.GetEncoding(1251).GetBytes("Город;Сумма\nМосква;\"1;5\"\nКазань;7\n");
    }

    private static byte[] GreyPng(int width, int height)
    {
        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

        byte[] header = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(header, width);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), height);
        header[8] = 8;
        Chunk(png, "IHDR", header);

        using var pixels = new MemoryStream();

        using (var deflated = new ZLibStream(pixels, CompressionLevel.Fastest, leaveOpen: true))
        {
            byte[] row = new byte[width + 1];

            for (int line = 0; line < height; line++)
                deflated.Write(row);
        }

        Chunk(png, "IDAT", pixels.ToArray());
        Chunk(png, "IEND", []);

        return png.ToArray();
    }

    private static void Chunk(Stream png, string type, byte[] data)
    {
        byte[] length = new byte[4];
        BinaryPrimitives.WriteInt32BigEndian(length, data.Length);
        byte[] typed = [.. Encoding.ASCII.GetBytes(type), .. data];
        byte[] crc = new byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crc, Crc32(typed));

        png.Write(length);
        png.Write(typed);
        png.Write(crc);
    }

    private static uint Crc32(byte[] bytes)
    {
        uint crc = 0xFFFFFFFF;

        foreach (byte value in bytes)
        {
            crc ^= value;

            for (int bit = 0; bit < 8; bit++)
                crc = (crc & 1) == 1 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }

        return ~crc;
    }

    /// <summary>The smallest PNG there is: one pixel.</summary>
    private static byte[] Png { get; } = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    /// <summary>A workbook of one sheet holding <paramref name="rows"/>, with <paramref name="styles"/> as its styles part when there is one.</summary>
    private static byte[] Workbook(S.Stylesheet? styles, params S.Row[] rows)
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();

            if (styles is not null)
                workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = styles;

            var sheet = workbook.AddNewPart<WorksheetPart>();
            sheet.Worksheet = new S.Worksheet(new S.SheetData(rows));

            workbook.Workbook.AppendChild(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1, Name = "Лист1" }));
        }

        return stream.ToArray();
    }

    private static W.Paragraph Paragraph(string text, W.ParagraphProperties? properties = null) =>
        properties is null
            ? new W.Paragraph(new W.Run(new W.Text(text)))
            : new W.Paragraph(properties, new W.Run(new W.Text(text)));

    private static W.TableRow Row(params string[] cells) =>
        new([.. cells.Select(cell => new W.TableCell(Paragraph(cell)))]);

    private static S.Row SheetRow(uint index, params S.Cell[] cells) =>
        new(cells) { RowIndex = index };

    private static S.Cell Shared(string reference, int index) =>
        new() { CellReference = reference, DataType = S.CellValues.SharedString, CellValue = new S.CellValue(index.ToString(System.Globalization.CultureInfo.InvariantCulture)) };

    private static S.Cell Number(string reference, string value) =>
        new() { CellReference = reference, CellValue = new S.CellValue(value) };

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
}
