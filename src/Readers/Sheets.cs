namespace Snail.Toolkit.Documents.Readers;

/// <summary>What every workbook format shares once its cells are read: one table a sheet, cut to its used columns.</summary>
internal static class Sheets
{
    /// <summary>The rows of a sheet as one table under the addresses they had; null for a sheet with nothing in it.</summary>
    /// <remarks>A sheet that starts at column A keeps its rows as they are: slicing every row from 0 copied each of them, 5 % of the Excel benchmark's allocations.</remarks>
    /// <param name="title">The sheet's name.</param>
    /// <param name="rows">The rows that hold something, each as wide as its last used column, from column A.</param>
    /// <param name="numbers">The sheet's number of each row.</param>
    public static TableBlock? Table(string? title, List<string[]> rows, List<int> numbers)
    {
        if (rows.Count == 0)
            return null;

        int first = rows.Min(row => Array.FindIndex(row, value => value.Length > 0));
        IReadOnlyList<IReadOnlyList<string>> used = first == 0 ? rows : [.. rows.Select(row => (IReadOnlyList<string>)row[first..])];
        var addresses = new CellAddresses(numbers, first);

        return new TableBlock(title, used) { Addresses = addresses };
    }
}
