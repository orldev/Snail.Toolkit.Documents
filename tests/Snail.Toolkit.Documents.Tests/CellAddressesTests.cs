namespace Snail.Toolkit.Documents.Tests;

/// <summary>A cell named the way the spreadsheet names it.</summary>
public sealed class CellAddressesTests
{
    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(701, "ZZ")]
    [InlineData(16_383, "XFD")]
    public void Letters_Column_IsNamedAsTheSpreadsheetNamesIt(int column, string letters) =>
        Assert.Equal(letters, CellAddresses.Letters(column));

    [Fact]
    public void Of_CellOfATableStartingAtC10_IsNamedFromTheSheet() =>
        Assert.Equal("D12", new CellAddresses([10, 12], FirstColumn: 2).At(1, 1));
}
