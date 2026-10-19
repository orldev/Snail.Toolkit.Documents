
namespace Snail.Toolkit.Documents.Tests;

public sealed class MarkdownTests
{
    /// <summary>The outline's positions are passed back as from; a position that missed would start in the wrong page.</summary>
    [Fact]
    public void Of_PartsHeadingsAndTables_PointAtWhereTheyStartInTheText()
    {
        var document = Markdown.Of(DocumentKind.Pdf,
        [
            new PartStart(PartKind.Page, 1, null),
            new Heading(2, "Итоги"),
            new TableBlock("t", [["a"], ["1"]]),
            new PartStart(PartKind.Page, 2, "Вторая")
        ]);

        Assert.All(document.Parts, part => Assert.StartsWith($"## Page {part.Number}", document.Text[part.Start..], StringComparison.Ordinal));
        Assert.StartsWith("## Итоги", document.Text[document.Headings[0].Start..], StringComparison.Ordinal);
        Assert.StartsWith("| a |", document.Text[document.Tables[0].Start..], StringComparison.Ordinal);
        Assert.Equal(document.Text.Length, document.Parts.Sum(part => part.Length));
        Assert.Equal(1, document.Tables[0].Part);
    }

    /// <summary>A bar inside a cell is a column border to Markdown unless it is escaped; a short row would shift the next column left.</summary>
    [Fact]
    public void Grid_CellWithABarAndARowShorterThanTheHeader_KeepsEveryColumnInPlace() =>
        Assert.Equal("| a | b |\n| --- | --- |\n| x\\|y |  |", Markdown.Grid([["a", "b"], ["x|y"]]));

    [Fact]
    public void Of_BlocksThatSayNothing_LeaveNoBlankLines()
    {
        var document = Markdown.Of(DocumentKind.Text, [new Paragraph("  "), new Paragraph("Текст"), new TableBlock(null, [])]);

        Assert.Equal("Текст", document.Text);
        Assert.Empty(document.Tables);
    }

    /// <summary>A block a newer adapter invents and this writer does not know is left out, not written as its type name.</summary>
    [Fact]
    public void Of_ABlockOfAKindNobodyWrites_LeavesItOut()
    {
        var document = Markdown.Of(DocumentKind.Text, [new Paragraph("Текст"), new Unknown()]);

        Assert.Equal("Текст", document.Text);
    }

    [Fact]
    public void Grid_RowsWithNoCells_IsNoTable() =>
        Assert.Equal(string.Empty, Markdown.Grid([[], []]));

    /// <summary>A file that unpacks to more text than a document may is refused while it is written, not after.</summary>
    [Fact]
    public void Of_BlocksPastTheCharacterBudget_IsRefusedAsTooLarge() =>
        Assert.Throws<DocumentTooLargeException>(() => Markdown.Of(DocumentKind.Text,
        [
            new Paragraph(new string('x', ReadingLimits.MostDocumentCharacters / 2)),
            new Paragraph(new string('y', ReadingLimits.MostDocumentCharacters / 2))
        ]));

    /// <summary>One wide row makes every row as wide; padded, a table of a few thousand cells becomes millions.</summary>
    [Fact]
    public void Grid_RowsPaddedPastTheCellBudget_IsRefusedAsTooLarge()
    {
        IReadOnlyList<string> wide = [.. Enumerable.Repeat("x", 1_001)];
        IReadOnlyList<IReadOnlyList<string>> rows = [wide, .. Enumerable.Repeat<IReadOnlyList<string>>(["x"], ReadingLimits.MostCells / 1_001)];

        Assert.Throws<DocumentTooLargeException>(() => Markdown.Grid(rows));
    }

    private sealed record Unknown : Block;
}
