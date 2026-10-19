using System.Text.RegularExpressions;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>A Markdown file: kept as written, with its headings, lists, tables and code recognised for the outline.</summary>
public sealed partial class MarkdownReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Markdown;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".md", ".markdown"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        string[] lines = Decoding.Text(content).ReplaceLineEndings("\n").Split('\n');
        List<Block> blocks = [];
        List<string> paragraph = [];

        for (int index = 0; index < lines.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string line = lines[index];

            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
            {
                Flush(blocks, paragraph);
                index = Fenced(lines, index, blocks);
                continue;
            }

            if (IsTableAt(lines, index))
            {
                Flush(blocks, paragraph);
                index = Tabled(lines, index, blocks);
                continue;
            }

            if (HeadingLine().Match(line) is { Success: true } heading)
            {
                Flush(blocks, paragraph);
                blocks.Add(new Heading(heading.Groups["marks"].Length, heading.Groups["text"].Value.Trim()));
                continue;
            }

            if (ItemLine().Match(line) is { Success: true } item)
            {
                Flush(blocks, paragraph);
                blocks.Add(new ListItem(item.Groups["indent"].Length / 2, item.Groups["text"].Value.Trim(), char.IsDigit(item.Groups["mark"].Value[0])));
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                Flush(blocks, paragraph);
                continue;
            }

            paragraph.Add(line);
        }

        Flush(blocks, paragraph);

        return blocks;
    }

    private static void Flush(List<Block> blocks, List<string> paragraph)
    {
        if (paragraph.Count > 0)
            blocks.Add(new Paragraph(string.Join('\n', paragraph)));

        paragraph.Clear();
    }

    private static int Fenced(string[] lines, int start, List<Block> blocks)
    {
        string language = lines[start].Trim()[3..].Trim();
        int end = start + 1;

        while (end < lines.Length && !lines[end].TrimStart().StartsWith("```", StringComparison.Ordinal))
            end++;

        blocks.Add(new CodeBlock(language, string.Join('\n', lines[(start + 1)..Math.Min(end, lines.Length)])));

        return end;
    }

    private static bool IsTableAt(string[] lines, int index) =>
        index + 1 < lines.Length && IsRow(lines[index]) && Separator().IsMatch(lines[index + 1]);

    private static int Tabled(string[] lines, int start, List<Block> blocks)
    {
        List<IReadOnlyList<string>> rows = [Cells(lines[start])];
        int index = start + 2;

        while (index < lines.Length && IsRow(lines[index]))
            rows.Add(Cells(lines[index++]));

        blocks.Add(new TableBlock(null, rows));

        return index - 1;
    }

    private static bool IsRow(string line) => line.TrimStart().StartsWith('|');

    private static string[] Cells(string line) =>
    [
        .. UnescapedBar().Split(line.Trim().Trim('|'))
            .Select(cell => cell.Replace("\\|", "|", StringComparison.Ordinal).Trim())
    ];

    [GeneratedRegex(@"^(?<marks>#{1,6})\s+(?<text>.+?)\s*#*\s*$", RegexOptions.CultureInvariant)]
    private static partial Regex HeadingLine();

    [GeneratedRegex(@"^(?<indent>\s*)(?<mark>[-*+]|\d+[.)])\s+(?<text>.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex ItemLine();

    /// <summary>The line under a table's header: dashes between bars, colons for alignment.</summary>
    /// <remarks>
    /// Without backtracking: the pattern opens with two runs of spaces around an optional bar, and the backtracking engine
    /// tried every split of them — a table row followed by 40 000 spaces took 22 s. The non-backtracking engine reads
    /// 400 000 spaces in 1 ms and accepts the same separators.
    /// </remarks>
    [GeneratedRegex(@"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$", RegexOptions.CultureInvariant | RegexOptions.NonBacktracking)]
    private static partial Regex Separator();

    [GeneratedRegex(@"(?<!\\)\|", RegexOptions.CultureInvariant)]
    private static partial Regex UnescapedBar();
}
