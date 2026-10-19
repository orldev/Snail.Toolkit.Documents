using System.Text;
using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents.Tests.Readers;

/// <summary>Every format, read from a real file into the same Markdown, outline and tables.</summary>
public sealed class FormatReaderTests
{
    /// <summary>A Russian Word saves «Heading 1» under the id «1»; read by id, the report would have no headings at all.</summary>
    [Fact]
    public async Task Word_HeadingStyledByARussianWord_IsAHeadingWithItsListAndTable()
    {
        var document = await ReadAsync(new WordReader(), Samples.Word());

        Assert.Equal([new Section(1, "Отчёт за неделю", 0)], document.Headings);
        Assert.Contains("# Отчёт за неделю\n\nПервый абзац отчёта.\n\n- пункт списка", document.Text, StringComparison.Ordinal);
        Assert.Equal([["Месяц", "Сумма"], ["Январь", "100"]], Rows(Assert.Single(document.Tables)));
    }

    /// <summary>A heading is known three ways besides a styled name; a report that used only one of them would come out flat.</summary>
    [Fact]
    public async Task Word_HeadingsByTitleOutlineAndStyleLevel_AreHeadingsOfTheirLevels()
    {
        var document = await ReadAsync(new WordReader(), Samples.WordOfEveryShape());

        Assert.Equal([(1, "Заголовок документа"), (2, "Раздел"), (3, "Подраздел")], document.Headings.Select(heading => (heading.Level, heading.Title)));
    }

    /// <summary>Text inside a content control is the document's text too; a reader that skipped it would lose whole forms.</summary>
    [Fact]
    public async Task Word_ContentControlTabsBreaksAndAnEmptyParagraph_ReadAsTheyLook()
    {
        var document = await ReadAsync(new WordReader(), Samples.WordOfEveryShape());

        Assert.Contains("Текст в элементе управления\n\nКолонка значение\nвторая строка\n\n1. пункт без определения списка", document.Text, StringComparison.Ordinal);
    }

    /// <summary>A date is a number in the file; only the cell's format says it is one.</summary>
    [Fact]
    public async Task Excel_SheetWithADateAndANoisyNumber_ReadsThemAsExcelShowsThem()
    {
        var document = await ReadAsync(new ExcelReader(), Samples.Excel());

        Assert.Equal(["Выручка", "Пусто"], document.Parts.Select(part => part.Title));
        Assert.Equal([["Дата", "Сумма"], ["2026-10-10", "0.3"], [string.Empty, "12"]], Rows(Assert.Single(document.Tables)));
        Assert.Equal(1, document.Tables[0].Part);
        Assert.StartsWith("## Sheet 1: Выручка\n\n|  | A | B |\n| --- | --- | --- |\n| 1 | Дата | Сумма |", document.Text, StringComparison.Ordinal);
    }

    /// <summary>A sheet that skips a row still names every cell where it stands: the third row of the table is row 4 of the sheet.</summary>
    [Fact]
    public async Task Excel_SheetThatSkipsARow_NamesEveryCellWhereItStands()
    {
        var document = await ReadAsync(new ExcelReader(), Samples.Excel());

        var addresses = document.Tables[0].Addresses!;

        Assert.Equal("B4", addresses.Of(2, 1));
        Assert.Contains("| 4 |  | 12 |", document.Text, StringComparison.Ordinal);
    }

    /// <summary>An older .xls reads as the same sheet a .xlsx would: dates by their format, numbers as Excel shows them, cells where they stand.</summary>
    [Fact]
    public async Task Xls_WorkbookOfTheOlderFormat_ReadsLikeAWorkbookOfTheNewer()
    {
        var document = await ReadAsync(new XlsReader(), Samples.OldExcel());

        var table = Assert.Single(document.Tables);

        Assert.Equal(["Счета", "Пусто"], document.Parts.Select(part => part.Title));
        Assert.Equal([["Номер", "Дата", "Сумма"], ["INV-7", "2026-10-14", "118000.5"], ["INV-8", string.Empty, "42"]], Rows(table));
        Assert.Equal("B2", table.Addresses!.Of(0, 0));
        Assert.Equal("D5", table.Addresses.Of(2, 2));
    }

    [Fact]
    public async Task Xls_BytesThatAreNotAWorkbook_AreCalledDamaged() =>
        await Assert.ThrowsAsync<InvalidDataException>(() => new XlsReader().ReadAsync(new MemoryStream(Samples.Csv())));

    /// <summary>Every cell reads as Excel shows it: a boolean as TRUE, an error as itself, a date by a format the workbook defines.</summary>
    [Fact]
    public async Task Excel_CellOfEveryKind_ReadsAsExcelShowsIt()
    {
        var document = await ReadAsync(new ExcelReader(), Samples.ExcelOfEveryCell());

        Assert.Equal(
            [["Строка", "TRUE", "FALSE", "#DIV/0!"], ["n/a", "2026-10-10", "12.5", "3000000"], ["без адреса"]],
            Rows(Assert.Single(document.Tables)));
    }

    /// <summary>Without a styles part nothing says a number is a date, so it is read as the number it is.</summary>
    [Fact]
    public async Task Excel_WorkbookWithoutStyles_ReadsANumberAsANumber()
    {
        var document = await ReadAsync(new ExcelReader(), Samples.ExcelWithoutStyles());

        Assert.Equal([["46305"]], Rows(Assert.Single(document.Tables)));
    }

    [Fact]
    public async Task PowerPoint_Slide_IsTitledAndCarriesItsPointsTableAndNotes()
    {
        var document = await ReadAsync(new PowerPointReader(), Samples.Slides());

        Assert.Equal("План работ", Assert.Single(document.Parts).Title);
        Assert.Contains("- Первое\n  - Второе", document.Text, StringComparison.Ordinal);
        Assert.Equal([["Этап", "Срок"], ["Сбор", "май"]], Rows(Assert.Single(document.Tables)));
        Assert.Contains("Notes: Сказать про сроки", document.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("- План работ", document.Text, StringComparison.Ordinal);
    }

    /// <summary>A slide keeps its number when the one before it cannot be read, and a text box in a group is read with its line break.</summary>
    [Fact]
    public async Task PowerPoint_UnlinkedSlideAndGroupedTextBox_KeepTheNumberingAndTheText()
    {
        var document = await ReadAsync(new PowerPointReader(), Samples.SlidesOfEveryShape());

        var slide = Assert.Single(document.Parts);
        Assert.Equal((2, "Схема"), (slide.Number, slide.Title));
        Assert.EndsWith("## Slide 2: Схема\n\nПервая\nвторая", document.Text, StringComparison.Ordinal);
    }

    /// <summary>A page with a picture and no text is a scan, said so — not an empty page passed over in silence.</summary>
    [Fact]
    public async Task Pdf_PageOfTextAndPageOfPicture_ReadsTheTextAndMarksTheScan()
    {
        var document = await ReadAsync(new PdfReader(), Samples.Pdf());

        Assert.Equal([false, true], document.Parts.Select(part => part.IsScanned));
        Assert.Contains("Quarterly report", document.Text, StringComparison.Ordinal);
        Assert.Contains("## Page 2\n\n*This page is a scan", document.Text, StringComparison.Ordinal);
    }

    [Fact]
    public Task Pdf_ThatIsNotAPdf_IsDamagedData() =>
        Assert.ThrowsAsync<InvalidDataException>(() => new PdfReader().ReadAsync(new MemoryStream("not a pdf"u8.ToArray())));

    /// <summary>A locked PDF says so, rather than being called damaged and asked for again.</summary>
    [Fact]
    public async Task Pdf_LockedByAPassword_IsRefusedAsLocked()
    {
        var refused = await Assert.ThrowsAsync<DocumentLockedException>(() => new PdfReader().ReadAsync(new MemoryStream(Samples.LockedPdf())));

        Assert.Equal("the PDF is protected by a password", refused.Message);
    }

    [Fact]
    public Task Word_ThatIsNotAPackage_IsDamagedData() =>
        Assert.ThrowsAsync<InvalidDataException>(() => new WordReader().ReadAsync(new MemoryStream("not a zip"u8.ToArray())));

    /// <summary>Read as UTF-8, a Windows-1251 CSV turns every Cyrillic letter into noise; a quoted cell keeps its separator.</summary>
    [Fact]
    public async Task Csv_SavedByARussianExcel_ReadsItsLettersSeparatorAndQuotedCell()
    {
        var document = await ReadAsync(new DelimitedReader(), Samples.Csv());

        Assert.Equal([["Город", "Сумма"], ["Москва", "1;5"], ["Казань", "7"]], Rows(Assert.Single(document.Tables)));
    }

    /// <summary>A separator inside a quoted header is not the separator, and a doubled quote inside a cell is one quote.</summary>
    [Fact]
    public async Task Csv_QuotedHeaderAndDoubledQuotes_KeepTheSeparatorAndTheQuotes()
    {
        var document = await ReadAsync(new DelimitedReader(), Encoding.UTF8.GetBytes("\"Название; полное\",Цена\n\"Торт \"\"Наполеон\"\"\",300\n"));

        Assert.Equal([["Название; полное", "Цена"], ["Торт \"Наполеон\"", "300"]], Rows(Assert.Single(document.Tables)));
    }

    /// <summary>A byte-order mark is believed: read past it as UTF-8, the mark would land in the text and UTF-16 would be noise.</summary>
    [Theory]
    [InlineData("utf-8")]
    [InlineData("utf-16")]
    [InlineData("utf-16BE")]
    public async Task Text_FileWithAByteOrderMark_IsReadInTheEncodingItMarks(string encoding)
    {
        var marked = Encoding.GetEncoding(encoding);

        var document = await ReadAsync(new PlainTextReader(), [.. marked.GetPreamble(), .. marked.GetBytes("Привет")]);

        Assert.Equal("Привет", document.Text);
    }

    [Fact]
    public async Task Json_ListOfRecords_IsATable()
    {
        var document = await ReadAsync(new JsonReader(), """[{"name":"a","n":1},{"name":"b","extra":true}]"""u8.ToArray());

        Assert.Equal([["name", "n", "extra"], ["a", "1", string.Empty], ["b", string.Empty, "true"]], Rows(Assert.Single(document.Tables)));
    }

    [Fact]
    public async Task Json_Tree_IsKeptAsIndentedJson()
    {
        var document = await ReadAsync(new JsonReader(), """{"a":{"b":[1,2]}}"""u8.ToArray());

        Assert.StartsWith("```json\n{\n  \"a\"", document.Text, StringComparison.Ordinal);
        Assert.Empty(document.Tables);
    }

    [Fact]
    public Task Json_ThatIsNotJson_IsDamagedData() =>
        Assert.ThrowsAsync<InvalidDataException>(() => new JsonReader().ReadAsync(new MemoryStream("{oops"u8.ToArray())));

    /// <summary>Most pages put their words loose in a div; a reader that waited for p would find them empty.</summary>
    [Fact]
    public async Task Html_PageWithMenuLooseTextNestedListAndTable_ReadsTheContentOnly()
    {
        var document = await ReadAsync(new HtmlReader(), Encoding.UTF8.GetBytes(
            """
            <html><head><title>Цены</title><style>p{}</style></head><body>
            <nav><a href="/">Главная</a></nav>
            <div>Текст без абзаца, <b>с выделением</b>.</div>
            <ul><li>Один<ul><li>Вложенный</li></ul></li><li>Два</li></ul>
            <table><tr><th>Товар</th><th>Цена</th></tr><tr><td>Хлеб</td><td>50</td></tr></table>
            <script>alert(1)</script>
            </body></html>
            """));

        Assert.Equal("Цены", Assert.Single(document.Headings).Title);
        Assert.Contains("Текст без абзаца, с выделением.", document.Text, StringComparison.Ordinal);
        Assert.Contains("- Один\n  - Вложенный\n- Два", document.Text, StringComparison.Ordinal);
        Assert.Equal([["Товар", "Цена"], ["Хлеб", "50"]], Rows(Assert.Single(document.Tables)));
        Assert.DoesNotContain("Главная", document.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("alert", document.Text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Html_HeadingsParagraphsQuoteAndPreformattedText_KeepTheirShape()
    {
        var document = await ReadAsync(new HtmlReader(), Encoding.UTF8.GetBytes(
            """
            <html><body>
            <h2>Раздел</h2>
            <p>Абзац   текста.</p>
            <blockquote>Цитата</blockquote>
            <pre>var  x = 1;</pre>
            </body></html>
            """));

        Assert.Equal([(2, "Раздел")], document.Headings.Select(heading => (heading.Level, heading.Title)));
        Assert.Contains("## Раздел\n\nАбзац текста.\n\nЦитата\n\n```\nvar  x = 1;\n```", document.Text, StringComparison.Ordinal);
    }

    /// <summary>Fifty thousand nested elements keep the parser busy for some ninety seconds; it stops when it is told to.</summary>
    [Fact]
    public async Task Html_PageNestedFiftyThousandDeep_StopsWhenItsTokenIsCancelled()
    {
        byte[] page = Encoding.UTF8.GetBytes($"<html><body>{string.Concat(Enumerable.Repeat("<div>", 50_000))}x</body></html>");
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(1));

        var reading = new HtmlReader().ReadAsync(new MemoryStream(page), deadline.Token);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reading.WaitAsync(TimeSpan.FromSeconds(15)));
    }

    [Fact]
    public async Task Markdown_File_KeepsItsHeadingsListTableAndCode()
    {
        var document = await ReadAsync(new MarkdownReader(), Encoding.UTF8.GetBytes(
            """
            # Заголовок

            Абзац.

            - пункт

            | a | b \| c |
            |---|---|
            | 1 | 2 |

            ```cs
            var x = 1;
            ```
            """));

        Assert.Equal(["Заголовок"], document.Headings.Select(heading => heading.Title));
        Assert.Equal([["a", "b | c"], ["1", "2"]], Rows(Assert.Single(document.Tables)));
        Assert.Contains("```cs\nvar x = 1;\n```", document.Text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Text_File_KeepsItsParagraphs()
    {
        var document = await ReadAsync(new PlainTextReader(), Encoding.UTF8.GetBytes("Первый\nабзац.\n\nВторой."));

        Assert.Equal("Первый\nабзац.\n\nВторой.", document.Text);
    }

    private static async Task<Document> ReadAsync(IFormatReader reader, byte[] bytes) =>
        Markdown.Of(reader.Format, await reader.ReadAsync(new MemoryStream(bytes)));

    private static string[][] Rows(Table table) => [.. table.Rows.Select(row => row.ToArray())];
}
