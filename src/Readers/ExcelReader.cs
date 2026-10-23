using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>An Excel workbook: every sheet a part, its used cells one table.</summary>
/// <remarks>
/// Cells keep the value the workbook last computed, so a formula reads as its result. A date is stored as a
/// number and known only by the format of its cell; read without the format, «2026-10-10» would reach the model
/// as 46305. Numbers are written to 15 significant digits, as Excel shows them, so a stored 0.30000000000000004
/// does not reach the model as a value nobody typed.
/// </remarks>
public sealed partial class ExcelReader : IFormatReader
{
    /// <summary>The last column a sheet has, XFD, counted from one.</summary>
    private const int LastColumn = 16_384;

    /// <summary>Number formats, and cell formats, a workbook's styles may define before the file is refused.</summary>
    /// <remarks>
    /// Excel itself keeps a workbook to some sixty-four thousand cell formats and a few hundred number formats; a million of each is
    /// fifteen times that and costs a few megabytes as indexes. Five million formats, a hundred megabytes the zip limits let through,
    /// stood at 2 GB as a tree.
    /// </remarks>
    private const int MostFormats = 1_000_000;

    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Excel;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".xlsx", ".xlsm"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        using var book = OfficePackage.Opened(content, package => SpreadsheetDocument.Open(package, isEditable: false));

        var workbook = book.WorkbookPart ?? throw new InvalidDataException("the file holds no workbook");
        var shared = SharedStrings(workbook.SharedStringTablePart, cancellationToken);
        var dates = DateStyles(workbook.WorkbookStylesPart, cancellationToken);
        var budget = new CharacterBudget();
        List<Block> blocks = [];
        int number = 0;

        foreach (var sheet in Listed(workbook, cancellationToken))
        {
            number++;
            var start = new PartStart(PartKind.Sheet, number, sheet.Name?.Value);
            budget.Count(Markdown.PartLine(start));
            blocks.Add(start);

            if (sheet.Id?.Value is not { } id || !workbook.TryGetPartById(id, out var part))
                throw new InvalidDataException($"the workbook lists the sheet '{sheet.Name?.Value}' under a relationship it does not hold");

            if (part is WorksheetPart worksheet && Table(sheet.Name?.Value, worksheet, shared, dates, budget, cancellationToken) is { } table)
                blocks.Add(table);
        }

        return blocks;
    }

    /// <summary>The sheets the workbook lists, in order, one entry of the list built at a time.</summary>
    /// <remarks>
    /// The list is the workbook's table of contents, a few lines in a real file; two million entries, a hundred megabytes the zip limits
    /// let through, stood at 1.1 GB as a tree before the first sheet was looked up. Read entry by entry, every sheet's heading is counted
    /// against the budget of characters as it is met, so a list too long for one document is refused at the budget.
    /// </remarks>
    private static IEnumerable<Sheet> Listed(WorkbookPart workbook, CancellationToken cancellationToken)
    {
        using var reader = OpenXmlReader.Create(workbook);

        while (reader.Read())
        {
            if (!reader.IsStartElement || reader.ElementType != typeof(Sheet))
                continue;

            cancellationToken.ThrowIfCancellationRequested();

            yield return (Sheet)reader.LoadCurrentElement()!;
        }
    }

    /// <summary>The shared strings of the workbook, read one at a time rather than as a tree, and refused past the budget of characters.</summary>
    /// <remarks>Excel writes only the strings its cells use, so a table of them past what one document may hold is a document that cannot be read.</remarks>
    private static IReadOnlyList<string> SharedStrings(SharedStringTablePart? part, CancellationToken cancellationToken)
    {
        if (part is null)
            return [];

        List<string> strings = [];
        long characters = 0;
        using var reader = OpenXmlReader.Create(part);

        while (reader.Read())
        {
            if (reader.ElementType != typeof(SharedStringItem) || !reader.IsStartElement)
                continue;

            cancellationToken.ThrowIfCancellationRequested();

            string text = reader.LoadCurrentElement()?.InnerText ?? string.Empty;
            characters += text.Length;

            if (characters > ReadingLimits.MostDocumentCharacters)
                throw new DocumentTooLargeException($"the workbook's shared strings alone come to more than {ReadingLimits.MostDocumentCharacters} characters");

            strings.Add(text);
        }

        return strings;
    }

    /// <summary>The used cells of a sheet as one table, each row under its number on the sheet; null for a sheet with none.</summary>
    /// <remarks>
    /// The sheet is read as the part reader's events — a row opens, a cell opens with its address, type and style, its value is
    /// text, the cell and the row close — and never loaded as a tree, not even a row or an inline string at a time: a tree of a
    /// million rows stood at 1.5 GB for 16 s before the budget of cells could refuse it, one step the token never interrupted, and
    /// a tree per row cost 5 % more than the whole on an ordinary ledger. Read this way, the rows kept are what the budgets of cells
    /// and characters allow and every row is a point the deadline can stop at.
    /// </remarks>
    private static TableBlock? Table(string? title, WorksheetPart part, IReadOnlyList<string> shared, IReadOnlySet<uint> dates, CharacterBudget budget, CancellationToken cancellationToken)
    {
        List<string[]> rows = [];
        List<int> numbers = [];
        List<(int Column, string Value)> cells = [];
        var cell = new CellReading();
        long padded = 0;
        using var reader = OpenXmlReader.Create(part);

        while (reader.Read())
        {
            if (reader.ElementType == typeof(Row) && reader.IsStartElement)
            {
                cancellationToken.ThrowIfCancellationRequested();
                cells.Clear();
                numbers.Add(RowNumber(reader, numbers));
            }
            else if (reader.ElementType == typeof(Cell) && reader.IsStartElement)
            {
                cell = CellReading.Opened(reader);
            }
            else if (reader.ElementType == typeof(CellValue) && reader.IsStartElement)
            {
                cell.Raw = reader.GetText();
            }
            else if (reader.ElementType == typeof(InlineString) && reader.IsStartElement)
            {
                cell.Inline = InlineText(reader, budget);
            }
            else if (reader.ElementType == typeof(Cell) && reader.IsEndElement)
            {
                if (Value(cell, shared, dates) is { Length: > 0 } value)
                {
                    budget.Count(value);
                    cells.Add((Column(cell.Reference) ?? cells.Count, value));
                }
            }
            else if (reader.ElementType == typeof(Row) && reader.IsEndElement)
            {
                padded += Kept(cells, rows, numbers);

                if (padded > ReadingLimits.MostCells)
                    throw new DocumentTooLargeException($"a sheet comes to more than {ReadingLimits.MostCells} cells once its rows are padded to their last column");
            }
        }

        return Sheets.Table(title, rows, numbers);
    }

    /// <summary>The text of the inline string the reader stands on, its pieces gathered to its end under the budget.</summary>
    private static string InlineText(OpenXmlReader reader, CharacterBudget budget)
    {
        int depth = reader.Depth;
        var text = new StringBuilder();

        while (reader.Read() && !(reader.IsEndElement && reader.Depth == depth))
        {
            if (!reader.IsStartElement || reader.ElementType != typeof(Text))
                continue;

            string piece = reader.GetText();
            budget.Count(piece);
            text.Append(piece);
        }

        return text.ToString();
    }

    /// <summary>Keeps the row just closed when it holds anything, padded to its last used column, and answers its width; a row of nothing is forgotten with its number.</summary>
    private static int Kept(List<(int Column, string Value)> cells, List<string[]> rows, List<int> numbers)
    {
        if (cells.Count == 0)
        {
            numbers.RemoveAt(numbers.Count - 1);

            return 0;
        }

        int width = cells.Max(cell => cell.Column) + 1;
        string[] filled = new string[width];
        Array.Fill(filled, string.Empty);

        foreach (var (column, value) in cells)
            filled[column] = value;

        rows.Add(filled);

        return width;
    }

    /// <summary>The sheet's number of the row the reader stands on: the one it states, or the one after the last row kept.</summary>
    private static int RowNumber(OpenXmlReader reader, List<int> numbers)
    {
        foreach (var attribute in reader.Attributes)
        {
            if (attribute.LocalName == "r" && uint.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint stated))
                return (int)stated;
        }

        return (numbers.Count > 0 ? numbers[^1] : 0) + 1;
    }

    private static string Value(CellReading cell, IReadOnlyList<string> shared, IReadOnlySet<uint> dates)
    {
        string raw = cell.Raw ?? string.Empty;

        if (cell.Type == "s")
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) && index < shared.Count ? shared[index] : string.Empty;

        if (cell.Type == "inlineStr")
            return cell.Inline ?? string.Empty;

        if (cell.Type == "b")
            return raw == "1" ? "TRUE" : "FALSE";

        if (cell.Type is not null && cell.Type != "n")
            return raw;

        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            return raw;

        return cell.Style is { } style && dates.Contains(style) ? Date(number) : number.ToString("G15", CultureInfo.InvariantCulture);
    }

    /// <summary>What a cell says about itself as it is read: its address, type and style from its attributes, its value and inline text from its children.</summary>
    private struct CellReading
    {
        public string? Reference;

        public string? Type;

        public uint? Style;

        public string? Raw;

        public string? Inline;

        public static CellReading Opened(OpenXmlReader reader)
        {
            var cell = new CellReading();

            foreach (var attribute in reader.Attributes)
            {
                switch (attribute.LocalName)
                {
                    case "r":
                        cell.Reference = attribute.Value;
                        break;
                    case "t":
                        cell.Type = attribute.Value;
                        break;
                    case "s":
                        cell.Style = uint.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint style) ? style : null;
                        break;
                }
            }

            return cell;
        }
    }

    private static string Date(double serial)
    {
        if (serial is < -657435 or > 2958465)
            return serial.ToString("G15", CultureInfo.InvariantCulture);

        var date = DateTime.FromOADate(serial);

        return date.TimeOfDay == TimeSpan.Zero
            ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
    }

    /// <summary>The zero-based column of a reference such as <c>AB12</c>; null without one.</summary>
    /// <remarks>
    /// A reference past XFD, the last column a sheet has, is no workbook Excel wrote: unbounded, <c>ZZZZZZZ1</c> overflowed
    /// the arithmetic into an array no process can allocate. Only ASCII letters are columns — any letter made a Cyrillic
    /// one column 1006.
    /// </remarks>
    internal static int? Column(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
            return null;

        int column = 0;

        foreach (char letter in reference.TakeWhile(char.IsAsciiLetter))
        {
            column = (column * 26) + (char.ToUpperInvariant(letter) - 'A' + 1);

            if (column > LastColumn)
                throw new InvalidDataException($"'{reference}' is past XFD, the last column a sheet has");
        }

        return column == 0 ? null : column - 1;
    }

    /// <summary>The indexes of the cell formats that show a date or a time, the styles read one format at a time.</summary>
    /// <remarks>
    /// Only the workbook's own number formats (<c>numFmts</c>) and cell formats (<c>cellXfs</c>) are read: the formats of cell
    /// styles and of conditional formatting carry the same elements and are not what a cell's style index points at. A number
    /// format defined twice is the later definition, as Excel reads it.
    /// </remarks>
    /// <exception cref="DocumentTooLargeException">The styles define more than <see cref="MostFormats"/> number or cell formats.</exception>
    private static IReadOnlySet<uint> DateStyles(WorkbookStylesPart? styles, CancellationToken cancellationToken)
    {
        HashSet<uint> dates = [];

        if (styles is null)
            return dates;

        Dictionary<uint, bool> isDateFormat = [];
        bool isInNumberFormats = false;
        bool isInCellFormats = false;
        uint index = 0;
        using var reader = OpenXmlReader.Create(styles);

        while (reader.Read())
        {
            if (reader.ElementType == typeof(NumberingFormats))
                isInNumberFormats = reader.IsStartElement;
            else if (reader.ElementType == typeof(CellFormats))
                isInCellFormats = reader.IsStartElement;
            else if (reader.IsStartElement && isInNumberFormats && reader.ElementType == typeof(NumberingFormat))
                Defined(isDateFormat, (NumberingFormat)reader.LoadCurrentElement()!, cancellationToken);
            else if (reader.IsStartElement && isInCellFormats && reader.ElementType == typeof(CellFormat))
                index = Formatted(dates, isDateFormat, (CellFormat)reader.LoadCurrentElement()!, index, cancellationToken);
        }

        return dates;
    }

    private static void Defined(Dictionary<uint, bool> isDateFormat, NumberingFormat format, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (format.NumberFormatId?.Value is not { } id)
            return;

        isDateFormat[id] = IsDateCode(format.FormatCode?.Value ?? string.Empty);

        if (isDateFormat.Count > MostFormats)
            throw new DocumentTooLargeException($"the workbook's styles define more than {MostFormats} number formats");
    }

    /// <summary>Notes the cell format at <paramref name="index"/> when it shows a date, and answers the index of the next one.</summary>
    private static uint Formatted(HashSet<uint> dates, Dictionary<uint, bool> isDateFormat, CellFormat format, uint index, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (index >= MostFormats)
            throw new DocumentTooLargeException($"the workbook's styles define more than {MostFormats} cell formats");

        uint id = format.NumberFormatId?.Value ?? 0;

        if (id is (>= 14 and <= 22) or (>= 45 and <= 47) || (isDateFormat.TryGetValue(id, out bool isDate) && isDate))
            dates.Add(index);

        return index + 1;
    }

    private static bool IsDateCode(string code) =>
        DateLetters().IsMatch(Literal().Replace(code, string.Empty));

    [GeneratedRegex(@"""[^""]*""|\[[^\]]*\]|\\.", RegexOptions.CultureInvariant)]
    private static partial Regex Literal();

    [GeneratedRegex("[ymdhs]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex DateLetters();
}
