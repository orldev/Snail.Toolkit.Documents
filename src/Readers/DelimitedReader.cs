using System.Text;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>Values separated by commas, semicolons or tabs: one table, its first line the header.</summary>
/// <remarks>
/// The separator of a .csv is guessed from its first line, counting outside quotes: a Russian Excel writes
/// semicolons into a file called .csv, because the comma is its decimal mark.
/// </remarks>
public sealed class DelimitedReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Csv;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".csv", ".tsv"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        string text = Decoding.Text(content).ReplaceLineEndings("\n");
        var rows = Rows(text, Separator(text), cancellationToken);

        return rows.Count == 0 ? [] : [new TableBlock(null, rows)];
    }

    /// <summary>The separator the first line uses most, outside quotes; a comma when it uses none.</summary>
    internal static char Separator(string text)
    {
        string first = text.Split('\n', 2)[0];
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
