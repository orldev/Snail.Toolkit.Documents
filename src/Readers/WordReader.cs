using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>A Word document: headings by their styles, lists by their numbering, paragraphs and tables in order.</summary>
/// <remarks>
/// A heading is known by the style's name, not its id: a Russian Word saves «Heading 1» under the id <c>1</c>
/// and a German one under <c>berschrift1</c>, while the name of a built-in style is always <c>heading 1</c>.
/// A style that sets an outline level counts as a heading of that level whatever it is called.
/// </remarks>
public sealed partial class WordReader : IFormatReader
{
    /// <summary>The ids of Word's built-in heading styles, with their levels.</summary>
    private static readonly IReadOnlyDictionary<string, int> BuiltInHeadings = new Dictionary<string, int>(StringComparer.Ordinal)
    {
        ["Title"] = 1,
        ["Heading1"] = 1,
        ["Heading2"] = 2,
        ["Heading3"] = 3,
        ["Heading4"] = 4,
        ["Heading5"] = 5,
        ["Heading6"] = 6
    };

    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Word;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".docx", ".docm"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        using var word = OfficePackage.Opened(content, package => WordprocessingDocument.Open(package, isEditable: false));

        var main = word.MainDocumentPart ?? throw new InvalidDataException("the file holds no document");
        var body = main.Document?.Body ?? throw new InvalidDataException("the document has no body");
        var levels = HeadingLevels(main.StyleDefinitionsPart);
        var bullets = BulletLevels(main.NumberingDefinitionsPart);

        return [.. Blocks(body, levels, bullets, cancellationToken)];
    }

    private static IEnumerable<Block> Blocks(OpenXmlElement container, IReadOnlyDictionary<string, int> levels, IReadOnlySet<(int, int)> bullets, CancellationToken cancellationToken)
    {
        foreach (var element in container.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();

            switch (element)
            {
                case W.Paragraph paragraph when Paragraphed(paragraph, levels, bullets) is { } block:
                    yield return block;
                    break;
                case W.Table table:
                    yield return new TableBlock(null, Rows(table));
                    break;
                case W.SdtBlock or W.CustomXmlBlock:
                    foreach (var inner in Blocks(element.Elements().LastOrDefault() ?? element, levels, bullets, cancellationToken))
                        yield return inner;
                    break;
            }
        }
    }

    private static Block? Paragraphed(W.Paragraph paragraph, IReadOnlyDictionary<string, int> levels, IReadOnlySet<(int, int)> bullets)
    {
        string text = Text(paragraph);
        if (string.IsNullOrWhiteSpace(text))
            return null;

        var properties = paragraph.ParagraphProperties;

        if (properties?.OutlineLevel?.Val?.Value is { } outline && outline < 6)
            return new Heading(outline + 1, text);

        if (properties?.ParagraphStyleId?.Val?.Value is { } style && levels.TryGetValue(style, out int level))
            return new Heading(level, text);

        if (properties?.NumberingProperties is { } numbering && numbering.NumberingId?.Val?.Value is { } list && list != 0)
        {
            int depth = numbering.NumberingLevelReference?.Val?.Value ?? 0;
            return new ListItem(depth, text, !bullets.Contains((list, depth)));
        }

        return new Paragraph(text);
    }

    private static string Text(OpenXmlElement element)
    {
        var text = new StringBuilder();

        foreach (var piece in element.Descendants())
        {
            switch (piece)
            {
                case W.Text written:
                    text.Append(written.Text);
                    break;
                case W.TabChar:
                    text.Append(' ');
                    break;
                case W.Break or W.CarriageReturn:
                    text.Append('\n');
                    break;
            }
        }

        return text.ToString().Trim();
    }

    private static IReadOnlyList<IReadOnlyList<string>> Rows(W.Table table) =>
        [.. table.Elements<W.TableRow>().Select(Cells).Where(row => row.Count > 0)];

    private static IReadOnlyList<string> Cells(W.TableRow row) =>
        [.. row.Elements<W.TableCell>().Select(CellText)];

    private static string CellText(W.TableCell cell) =>
        string.Join(' ', cell.Elements<W.Paragraph>().Select(Text).Where(line => line.Length > 0));

    /// <summary>The heading level of every style that is a heading, by style id.</summary>
    /// <remarks>
    /// The built-in ids count as headings until the file defines them otherwise: a document a program generated names the
    /// style <c>Heading1</c> and leaves the styles part out, and Word itself shows such a paragraph as a heading. A style the
    /// file does define is a heading only if its outline level or its name says so.
    /// </remarks>
    private static IReadOnlyDictionary<string, int> HeadingLevels(StyleDefinitionsPart? styles)
    {
        Dictionary<string, int> levels = new(BuiltInHeadings, StringComparer.Ordinal);

        foreach (var style in styles?.Styles?.Elements<W.Style>() ?? [])
        {
            if (style.StyleId?.Value is not { } id)
                continue;

            levels.Remove(id);

            if (style.StyleParagraphProperties?.OutlineLevel?.Val?.Value is { } outline && outline < 6)
                levels[id] = outline + 1;
            else if (style.StyleName?.Val?.Value is { } name && Level(name) is { } level)
                levels[id] = level;
        }

        return levels;
    }

    private static int? Level(string name)
    {
        if (name.Equals("title", StringComparison.OrdinalIgnoreCase))
            return 1;

        return HeadingName().Match(name) is { Success: true } match ? int.Parse(match.Groups["level"].Value, System.Globalization.CultureInfo.InvariantCulture) : null;
    }

    /// <summary>The list and level pairs whose numbering draws a bullet rather than a number.</summary>
    private static IReadOnlySet<(int, int)> BulletLevels(NumberingDefinitionsPart? numbering)
    {
        if (numbering?.Numbering is not { } definitions)
            return new HashSet<(int, int)>();

        var abstracts = definitions.Elements<W.AbstractNum>()
            .Where(definition => definition.AbstractNumberId?.Value is not null)
            .ToDictionary(definition => definition.AbstractNumberId!.Value, definition => definition);

        HashSet<(int, int)> bullets = [];

        foreach (var instance in definitions.Elements<W.NumberingInstance>())
        {
            if (instance.NumberID?.Value is not { } list || instance.AbstractNumId?.Val?.Value is not { } id || !abstracts.TryGetValue(id, out var definition))
                continue;

            foreach (var level in definition.Elements<W.Level>().Where(level => level.NumberingFormat?.Val?.Value == W.NumberFormatValues.Bullet))
                bullets.Add((list, level.LevelIndex?.Value ?? 0));
        }

        return bullets;
    }

    [GeneratedRegex(@"^heading\s+(?<level>[1-6])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex HeadingName();
}
