using System.Text;

namespace Snail.Toolkit.Documents;

/// <summary>Writes the blocks of a document as Markdown, and notes where each part, heading and table starts.</summary>
/// <remarks>
/// The text and its outline come from one pass, so a position the outline gives is the position the text has:
/// the model passes it back as <c>from</c>, and a second pass that counted differently would send it into the
/// middle of the wrong page.
/// </remarks>
public static class Markdown
{
    /// <summary>The document the blocks make.</summary>
    /// <exception cref="DocumentTooLargeException">The text, or one of its tables, comes to more than one document may.</exception>
    public static Document Of(DocumentKind format, IReadOnlyList<Block> blocks)
    {
        var text = new StringBuilder();
        List<(PartStart Start, int At)> parts = [];
        List<Section> headings = [];
        List<Table> tables = [];
        Block? previous = null;

        foreach (var block in blocks)
        {
            if (Written(block) is not { Length: > 0 } written)
                continue;

            string separator = Separator(text.Length, previous, block);

            if (text.Length + separator.Length + written.Length > ReadingLimits.MostDocumentCharacters)
                throw new DocumentTooLargeException($"its text comes to more than {ReadingLimits.MostDocumentCharacters} characters");

            text.Append(separator);

            int at = text.Length;

            switch (block)
            {
                case PartStart start:
                    parts.Add((start, at));
                    break;
                case Heading heading:
                    headings.Add(new Section(heading.Level, heading.Text, at));
                    break;
                case TableBlock table:
                    tables.Add(new Table(tables.Count + 1, table.Title, parts.Count > 0 ? parts[^1].Start.Number : null, at, table.Rows));
                    break;
            }

            text.Append(written);
            previous = block;
        }

        string whole = text.ToString();

        return new Document(format, whole, Measured(parts, whole.Length), headings, tables);
    }

    /// <summary>A block as Markdown; empty for one that says nothing.</summary>
    internal static string Written(Block block) => block switch
    {
        PartStart start => PartLine(start),
        Heading heading => string.IsNullOrWhiteSpace(heading.Text) ? string.Empty : $"{new string('#', Math.Clamp(heading.Level, 1, 6))} {Line(heading.Text)}",
        Paragraph paragraph => paragraph.Text.Trim(),
        ListItem item => string.IsNullOrWhiteSpace(item.Text) ? string.Empty : $"{new string(' ', Math.Clamp(item.Depth, 0, ReadingLimits.MostListDepth) * 2)}{(item.IsOrdered ? "1." : "-")} {Line(item.Text)}",
        TableBlock table => Grid(table.Rows),
        CodeBlock code => string.IsNullOrWhiteSpace(code.Text) ? string.Empty : $"```{code.Language}\n{code.Text.TrimEnd()}\n```",
        _ => string.Empty
    };

    /// <summary>The heading that opens a part: what it is, its number and title, and whether it is a scan.</summary>
    internal static string PartLine(PartStart start)
    {
        string name = $"## {start.Kind} {start.Number}{(string.IsNullOrWhiteSpace(start.Title) ? string.Empty : $": {Line(start.Title)}")}";

        return start.IsScanned ? $"{name}\n\n*This page is a scan with no text layer: its words are not in this text.*" : name;
    }

    /// <summary>Rows as a Markdown table, the first of them the header; every row as wide as the widest.</summary>
    /// <remarks>Refused past <see cref="ReadingLimits.MostCells"/>: padding every row to the widest is what turns a ragged table of a few rows into millions of empty cells.</remarks>
    internal static string Grid(IReadOnlyList<IReadOnlyList<string>> rows)
    {
        if (rows.Count == 0)
            return string.Empty;

        int columns = rows.Max(row => row.Count);
        if (columns == 0)
            return string.Empty;

        if ((long)rows.Count * columns > ReadingLimits.MostCells)
            throw new DocumentTooLargeException($"a table of {rows.Count} rows and {columns} columns comes to more than {ReadingLimits.MostCells} cells");

        var grid = new StringBuilder();

        for (int index = 0; index < rows.Count; index++)
        {
            var cells = Enumerable.Range(0, columns).Select(column => column < rows[index].Count ? Cell(rows[index][column]) : string.Empty);
            grid.Append($"| {string.Join(" | ", cells)} |");

            if (index == 0)
                grid.Append($"\n|{string.Concat(Enumerable.Repeat(" --- |", columns))}");

            if (index < rows.Count - 1)
                grid.Append('\n');
        }

        return grid.ToString();
    }

    private static string Cell(string text) =>
        Line(text).Replace("|", "\\|", StringComparison.Ordinal);

    private static string Line(string text) =>
        string.Join(' ', text.Split((char[])['\r', '\n', '\t'], StringSplitOptions.RemoveEmptyEntries).Select(piece => piece.Trim()).Where(piece => piece.Length > 0));

    /// <summary>What goes before a block: nothing at the start, a line break between list items, a blank line otherwise.</summary>
    private static string Separator(int written, Block? previous, Block block)
    {
        if (written == 0)
            return string.Empty;

        return IsListItemAfterListItem(previous, block) ? "\n" : "\n\n";
    }

    private static bool IsListItemAfterListItem(Block? previous, Block block) =>
        previous is ListItem && block is ListItem;

    private static IReadOnlyList<Part> Measured(List<(PartStart Start, int At)> parts, int total) =>
    [
        .. parts.Select((part, index) => new Part(
            part.Start.Kind,
            part.Start.Number,
            part.Start.Title,
            part.At,
            (index + 1 < parts.Count ? parts[index + 1].At : total) - part.At,
            part.Start.IsScanned))
    ];
}
