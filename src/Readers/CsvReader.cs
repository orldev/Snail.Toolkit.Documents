using System.Text;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>Values separated by commas or semicolons, a .csv: one table, its first line the header.</summary>
/// <remarks>
/// The separator of a .csv is guessed from its first line, counting outside quotes: a Russian Excel writes
/// semicolons into a file called .csv, because the comma is its decimal mark. A .tsv names its separator and is
/// read by <see cref="TsvReader"/>, since a reader never sees the file's name.
/// </remarks>
public sealed class CsvReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Kind => DocumentKind.Csv;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".csv"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, null, cancellationToken));

    /// <summary>The table of the text, split by <paramref name="separator"/>, or by the one its first line uses most when none is given.</summary>
    internal static IReadOnlyList<Block> Read(Stream content, char? separator, CancellationToken cancellationToken)
    {
        string text = Decoding.Text(content).ReplaceLineEndings("\n");
        var rows = Rows(text, separator ?? Separator(text), cancellationToken);

        return rows.Count == 0 ? [] : [new TableBlock(null, rows)];
    }

    /// <summary>The separator the first line uses most, outside quotes; a comma when it uses none.</summary>
    /// <remarks>The first line is a span of the text, not a split: splitting copied the rest of the file to read one line, 30 MB for a file of twelve million characters.</remarks>
    internal static char Separator(string text)
    {
        int end = text.IndexOf('\n', StringComparison.Ordinal);
        ReadOnlySpan<char> first = end < 0 ? text : text.AsSpan(0, end);
        bool isQuoted = false;
        Dictionary<char, int> counts = new() { [','] = 0, [';'] = 0, ['\t'] = 0 };

        foreach (char letter in first)
        {
            if (letter == '"')
                isQuoted = !isQuoted;
            else if (!isQuoted && counts.TryGetValue(letter, out int count))
                counts[letter] = count + 1;
        }

        var most = counts.MaxBy(count => count.Value);

        return most.Value == 0 ? ',' : most.Key;
    }

    /// <summary>The rows of the text, with quoted fields that hold separators, quotes and line breaks read as one cell.</summary>
    internal static IReadOnlyList<IReadOnlyList<string>> Rows(string text, char separator, CancellationToken cancellationToken = default)
    {
        List<IReadOnlyList<string>> rows = [];
        List<string> row = [];
        var cell = new StringBuilder();
        bool isQuoted = false;

        for (int index = 0; index < text.Length; index++)
        {
            char letter = text[index];

            if (isQuoted)
            {
                if (letter == '"' && index + 1 < text.Length && text[index + 1] == '"')
                {
                    cell.Append('"');
                    index++;
                }
                else if (letter == '"')
                {
                    isQuoted = false;
                }
                else
                {
                    cell.Append(letter);
                }

                continue;
            }

            if (letter == '"' && cell.Length == 0)
            {
                isQuoted = true;
            }
            else if (letter == separator)
            {
                row.Add(cell.ToString());
                cell.Clear();
            }
            else if (letter == '\n')
            {
                cancellationToken.ThrowIfCancellationRequested();
                row.Add(cell.ToString());
                cell.Clear();
                rows.Add([.. row]);
                row.Clear();
            }
            else
            {
                cell.Append(letter);
            }
        }

        if (cell.Length > 0 || row.Count > 0)
        {
            row.Add(cell.ToString());
            rows.Add([.. row]);
        }

        return [.. rows.Where(cells => cells.Any(value => value.Length > 0))];
    }
}
