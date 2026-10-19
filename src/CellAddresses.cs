namespace Snail.Toolkit.Documents;

/// <summary>Where the cells of a sheet's table stand on the sheet, so a cell is named the way the spreadsheet names it.</summary>
/// <param name="Rows">The sheet's number of each row of the table, from 1; a row with nothing in it is not in the table.</param>
/// <param name="FirstColumn">The sheet's column of the table's first column, from 0 for A.</param>
/// <remarks>
/// A model that quotes a value names the cell it took it from, and a person checks it there: a table that lost where it
/// stood on the sheet would send them to the wrong row as soon as the sheet skipped one.
/// </remarks>
public sealed record CellAddresses(IReadOnlyList<int> Rows, int FirstColumn)
{
    /// <summary>The address of the cell at a row and column of the table, such as <c>B7</c>.</summary>
    public string Of(int row, int column) => $"{Letters(FirstColumn + column)}{Rows[row]}";

    /// <summary>The letters of a column counted from 0: 0 is A, 25 is Z, 26 is AA.</summary>
    public static string Letters(int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column);

        string letters = string.Empty;

        for (int remaining = column + 1; remaining > 0; remaining = (remaining - 1) / 26)
            letters = $"{(char)('A' + ((remaining - 1) % 26))}{letters}";

        return letters;
    }
}
