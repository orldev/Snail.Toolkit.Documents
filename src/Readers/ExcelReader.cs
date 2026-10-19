using System.Globalization;
using System.Text.RegularExpressions;
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
        IReadOnlyList<string> shared = [.. workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(item => item.InnerText) ?? []];
        var dates = DateStyles(workbook.WorkbookStylesPart?.Stylesheet);
        List<Block> blocks = [];
        int number = 0;

        foreach (var sheet in workbook.Workbook?.Sheets?.Elements<Sheet>() ?? [])
        {
            number++;
            blocks.Add(new PartStart(PartKind.Sheet, number, sheet.Name?.Value));

            if (sheet.Id?.Value is { } id && workbook.GetPartById(id) is WorksheetPart part && Table(sheet.Name?.Value, part, shared, dates, cancellationToken) is { } table)
                blocks.Add(table);
        }

        return blocks;
    }

    /// <summary>The used cells of a sheet as one table, each row under its number on the sheet; null for a sheet with none.</summary>
    private static TableBlock? Table(string? title, WorksheetPart part, IReadOnlyList<string> shared, IReadOnlySet<uint> dates, CancellationToken cancellationToken)
    {
        List<string[]> rows = [];
        List<int> numbers = [];
        long padded = 0;

        foreach (var row in part.Worksheet?.Descendants<Row>() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();

            List<(int Column, string Value)> cells = [];

            foreach (var cell in row.Elements<Cell>())
            {
                if (Value(cell, shared, dates) is { Length: > 0 } value)
                    cells.Add((Column(cell.CellReference?.Value) ?? cells.Count, value));
            }

            if (cells.Count == 0)
                continue;

            int width = cells.Max(cell => cell.Column) + 1;
            padded += width;

            if (padded > ReadingLimits.MostCells)
                throw new DocumentTooLargeException($"a sheet comes to more than {ReadingLimits.MostCells} cells once its rows are padded to their last column");

            string[] filled = new string[width];
            Array.Fill(filled, string.Empty);

            foreach (var (column, value) in cells)
                filled[column] = value;

            rows.Add(filled);
            numbers.Add(row.RowIndex?.Value is { } index ? (int)index : (numbers.Count > 0 ? numbers[^1] : 0) + 1);
        }

        return Sheets.Table(title, rows, numbers);
    }

    private static string Value(Cell cell, IReadOnlyList<string> shared, IReadOnlySet<uint> dates)
    {
        string raw = cell.CellValue?.Text ?? string.Empty;

        if (cell.DataType?.Value == CellValues.SharedString)
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) && index < shared.Count ? shared[index] : string.Empty;

        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? string.Empty;

        if (cell.DataType?.Value == CellValues.Boolean)
            return raw == "1" ? "TRUE" : "FALSE";

        if (cell.DataType is not null && cell.DataType.Value != CellValues.Number)
            return raw;

        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            return raw;

        return cell.StyleIndex?.Value is { } style && dates.Contains(style) ? Date(number) : number.ToString("G15", CultureInfo.InvariantCulture);
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

    /// <summary>The indexes of the cell formats that show a date or a time.</summary>
    private static IReadOnlySet<uint> DateStyles(Stylesheet? styles)
    {
        if (styles?.CellFormats is not { } formats)
            return new HashSet<uint>();

        var custom = styles.NumberingFormats?.Elements<NumberingFormat>()
            .Where(format => format.NumberFormatId?.Value is not null)
            .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode?.Value ?? string.Empty) ?? [];

        HashSet<uint> dates = [];
        uint index = 0;

        foreach (var format in formats.Elements<CellFormat>())
        {
            uint id = format.NumberFormatId?.Value ?? 0;

            if (id is (>= 14 and <= 22) or (>= 45 and <= 47) || (custom.TryGetValue(id, out string? code) && IsDateCode(code)))
                dates.Add(index);

            index++;
        }

        return dates;
    }

    private static bool IsDateCode(string code) =>
        DateLetters().IsMatch(Literal().Replace(code, string.Empty));

    [GeneratedRegex(@"""[^""]*""|\[[^\]]*\]|\\.", RegexOptions.CultureInvariant)]
    private static partial Regex Literal();

    [GeneratedRegex("[ymdhs]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex DateLetters();
}
