using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>A JSON file: a list of flat records as a table, anything else as indented JSON.</summary>
/// <remarks>
/// A list of records is what an export usually is, and a table of it is what a model counts and compares in;
/// any other shape is kept as JSON, indented, because flattening a tree loses which value belongs to what.
/// </remarks>
public sealed class JsonReader : IFormatReader
{
    private static readonly JsonSerializerOptions Indented = new() { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Json;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".json"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        JsonNode? root;

        try
        {
            root = JsonNode.Parse(Decoding.Text(content));
        }
        catch (JsonException failure)
        {
            throw new InvalidDataException(failure.Message, failure);
        }

        if (root is JsonArray { Count: > 0 } records && records.All(IsFlatRecord))
            return [new TableBlock(null, Table(records))];

        return [new CodeBlock("json", root?.ToJsonString(Indented) ?? "null")];
    }

    private static bool IsFlatRecord(JsonNode? node) =>
        node is JsonObject record && record.All(field => field.Value is null or JsonValue);

    /// <summary>The records as rows under a header of every key any of them names.</summary>
    /// <remarks>
    /// The header is the union of the keys, so records that each name a key of their own make a grid as wide as it is
    /// long: two thousand records of one field came to 12 032 893 characters. Refused past <see cref="ReadingLimits.MostCells"/>.
    /// </remarks>
    private static IReadOnlyList<IReadOnlyList<string>> Table(JsonArray records)
    {
        List<string> columns = [.. records.Cast<JsonObject>().SelectMany(record => record.Select(field => field.Key)).Distinct(StringComparer.Ordinal)];
        long cells = (long)columns.Count * (records.Count + 1);

        if (cells > ReadingLimits.MostCells)
            throw new DocumentTooLargeException($"{records.Count} records under {columns.Count} keys come to more than {ReadingLimits.MostCells} cells");

        return
        [
            columns,
            .. records.Cast<JsonObject>().Select(record => (IReadOnlyList<string>)[.. columns.Select(column => Value(record[column]))])
        ];
    }

    private static string Value(JsonNode? value) => value switch
    {
        null => string.Empty,
        JsonValue text when text.GetValueKind() == JsonValueKind.String => text.GetValue<string>(),
        _ => value.ToJsonString()
    };
}
