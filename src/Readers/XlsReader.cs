using System.Globalization;
using System.Text;
using ExcelDataReader;
using ExcelDataReader.Exceptions;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>An Excel workbook of the older binary format, <c>.xls</c>: every sheet a part, its used cells one table.</summary>
/// <remarks>
/// Excel saved this format by default until 2007, and accounting systems still export it. It is read with ExcelDataReader,
/// the managed reader of the format, and written exactly as <see cref="ExcelReader"/> writes a workbook — dates by their
/// cell format, numbers to 15 significant digits — so the same sheet reads the same whichever format it was saved in.
/// The format's text is in the code page of the machine that wrote it, so the code pages are registered first.
/// </remarks>
public sealed class XlsReader : IFormatReader
{
    static XlsReader() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <inheritdoc/>
    public DocumentKind Kind => DocumentKind.Excel;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".xls"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        using var book = Opened(content);
        List<Block> blocks = [];
        int number = 0;

        do
        {
            number++;
            blocks.Add(new PartStart(PartKind.Sheet, number, book.Name));

            if (Table(book, cancellationToken) is { } table)
                blocks.Add(table);
        }
        while (book.NextResult());

        return blocks;
    }

    /// <summary>The used cells of the sheet the reader stands on, each row under its number on the sheet.</summary>
    private static TableBlock? Table(IExcelDataReader book, CancellationToken cancellationToken)
    {
        List<string[]> rows = [];
        List<int> numbers = [];
        long padded = 0;

        for (int row = 1; HasNextRow(book); row++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            string[] cells = [.. Enumerable.Range(0, book.FieldCount).Select(column => Value(book.GetValue(column)))];
            int width = Array.FindLastIndex(cells, value => value.Length > 0) + 1;

            if (width == 0)
                continue;

            padded += width;

            if (padded > ReadingLimits.MostCells)
                throw new DocumentTooLargeException($"a sheet comes to more than {ReadingLimits.MostCells} cells once its rows are padded to their last column");

            rows.Add(cells[..width]);
            numbers.Add(row);
        }

        return UsedRange.Table(book.Name, rows, numbers);
    }

    private static string Value(object? value) => value switch
    {
        null => string.Empty,
        DateTime date when date.TimeOfDay == TimeSpan.Zero => date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        DateTime date => date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
        double number => number.ToString("G15", CultureInfo.InvariantCulture),
        bool flag => flag ? "TRUE" : "FALSE",
        IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString() ?? string.Empty
    };

    /// <summary>The workbook opened, every way it can fail to open turned into the failure the readers share.</summary>
    private static IExcelDataReader Opened(Stream content)
    {
        try
        {
            return ExcelReaderFactory.CreateBinaryReader(content, new ExcelReaderConfiguration { LeaveOpen = true });
        }
        catch (InvalidPasswordException failure)
        {
            throw new DocumentLockedException("the workbook is protected by a password", failure);
        }
        catch (Exception failure) when (failure is ExcelReaderException or NotSupportedException or EndOfStreamException)
        {
            throw new InvalidDataException(failure.Message, failure);
        }
    }

    /// <summary>Whether there is a next row, moving to it; a record the format's reader cannot parse is a damaged file, not a crash.</summary>
    private static bool HasNextRow(IExcelDataReader book)
    {
        try
        {
            return book.Read();
        }
        catch (Exception failure) when (failure is ExcelReaderException or EndOfStreamException)
        {
            throw new InvalidDataException(failure.Message, failure);
        }
    }
}
