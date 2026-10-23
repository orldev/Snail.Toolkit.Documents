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
/// Nothing of the document is built as a tree but a paragraph's properties and one style or numbering at a time:
/// the body is read as the part reader's events, and the text is gathered run by run under the budget of characters,
/// so what the heap holds is the budget, not the file — a paragraph of fifty million letters stood at 337 MB as a tree,
/// a table of nine hundred thousand rows at 1 GB for 23 s.
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
        var shape = new Shape(HeadingLevels(main.StyleDefinitionsPart), BulletLevels(main.NumberingDefinitionsPart), new CharacterBudget(), cancellationToken);

        return [.. Blocks(main, shape)];
    }

    /// <summary>The blocks of the document's body: every paragraph and table the part reader meets, each read to its end as it is met.</summary>
    /// <remarks>A paragraph inside a content control arrives as its own event; one inside a table is read as part of the table.</remarks>
    private static IEnumerable<Block> Blocks(MainDocumentPart main, Shape shape)
    {
        using var reader = OpenXmlReader.Create(main);
        bool hasBody = false;

        while (reader.Read())
        {
            if (!reader.IsStartElement)
                continue;

            shape.Token.ThrowIfCancellationRequested();

            if (reader.ElementType == typeof(W.Body))
                hasBody = true;
            else if (reader.ElementType == typeof(W.Paragraph) && Paragraphed(reader, shape) is { } block)
                yield return block;
            else if (reader.ElementType == typeof(W.Table))
                yield return new TableBlock(null, Rows(reader, shape));
        }

        if (!hasBody)
            throw new InvalidDataException("the document has no body");
    }

    /// <summary>The paragraph the reader stands on, read to its end: its properties built, its text gathered run by run under the budget; null for one that says nothing.</summary>
    /// <remarks>
    /// Word writes a text box inside <c>mc:AlternateContent</c> twice: as a shape for programs that know shapes and, in
    /// <c>mc:Fallback</c>, as the older drawing for those that do not. A walk that entered both read every text box twice over;
    /// the fallback is passed over. The text of a box, a nested paragraph, is part of the paragraph that holds it.
    /// </remarks>
    private static Block? Paragraphed(OpenXmlReader reader, Shape shape)
    {
        int depth = reader.Depth;
        W.ParagraphProperties? properties = null;
        var text = new StringBuilder();
        int skipped = -1;

        while (reader.Read() && !IsEndAt(reader, depth))
        {
            if (skipped >= 0)
            {
                skipped = IsEndAt(reader, skipped) ? -1 : skipped;
                continue;
            }

            if (!reader.IsStartElement)
                continue;

            if (reader.ElementType == typeof(W.ParagraphProperties) && reader.Depth == depth + 1)
                properties = (W.ParagraphProperties)reader.LoadCurrentElement()!;
            else if (reader.ElementType == typeof(AlternateContentFallback))
                skipped = reader.Depth;
            else
                Append(reader, text, shape.Budget);
        }

        string written = text.ToString().Trim();

        return string.IsNullOrWhiteSpace(written) ? null : Classified(written, properties, shape);
    }

    private static void Append(OpenXmlReader reader, StringBuilder text, CharacterBudget budget)
    {
        if (reader.ElementType == typeof(W.Text))
        {
            string piece = reader.GetText();
            budget.Count(piece);
            text.Append(piece);
        }
        else if (reader.ElementType == typeof(W.TabChar))
        {
            text.Append(' ');
        }
        else if (reader.ElementType == typeof(W.Break) || reader.ElementType == typeof(W.CarriageReturn))
        {
            text.Append('\n');
        }
    }

    /// <summary>What a paragraph with this text and these properties is: a heading by its outline level or style, an item of a list by its numbering, a paragraph otherwise.</summary>
    private static Block Classified(string text, W.ParagraphProperties? properties, Shape shape)
    {
        if (properties?.OutlineLevel?.Val?.Value is { } outline && outline < 6)
            return new Heading(outline + 1, text);

        if (properties?.ParagraphStyleId?.Val?.Value is { } style && shape.Levels.TryGetValue(style, out int level))
            return new Heading(level, text);

        if (properties?.NumberingProperties is { } numbering && numbering.NumberingId?.Val?.Value is { } list && list != 0)
        {
            int depth = numbering.NumberingLevelReference?.Val?.Value ?? 0;

            return new ListItem(depth, text, !shape.Bullets.Contains((list, depth)));
        }

        return new Paragraph(text);
    }

    /// <summary>The rows of the table the reader stands on, read to its end, each cell flattened to one line.</summary>
    private static IReadOnlyList<IReadOnlyList<string>> Rows(OpenXmlReader reader, Shape shape)
    {
        int depth = reader.Depth;
        List<IReadOnlyList<string>> rows = [];
        List<string> cells = [];

        while (reader.Read() && !IsEndAt(reader, depth))
        {
            shape.Token.ThrowIfCancellationRequested();

            if (reader.IsStartElement && reader.ElementType == typeof(W.TableRow) && reader.Depth == depth + 1)
                cells = [];
            else if (reader.IsStartElement && reader.ElementType == typeof(W.TableCell) && reader.Depth == depth + 2)
                cells.Add(CellText(reader, shape));
            else if (reader.IsEndElement && reader.ElementType == typeof(W.TableRow) && reader.Depth == depth + 1 && cells.Count > 0)
                rows.Add(cells);
        }

        return rows;
    }

    /// <summary>Everything the cell the reader stands on holds, flattened to one line: its paragraphs, those of the content controls in it and those of a table nested in it.</summary>
    private static string CellText(OpenXmlReader reader, Shape shape)
    {
        int depth = reader.Depth;
        List<string> lines = [];

        while (reader.Read() && !IsEndAt(reader, depth))
        {
            if (reader.IsStartElement && reader.ElementType == typeof(W.Paragraph) && Paragraphed(reader, shape) is { } block && Line(block) is { Length: > 0 } line)
                lines.Add(line);
        }

        return string.Join(' ', lines);
    }

    private static string Line(Block block) => block switch
    {
        Heading heading => heading.Text,
        Paragraph paragraph => paragraph.Text,
        ListItem item => item.Text,
        _ => string.Empty
    };

    private static bool IsEndAt(OpenXmlReader reader, int depth) =>
        reader.IsEndElement && reader.Depth == depth;

    /// <summary>The heading level of every style that is a heading, by style id, one style built at a time.</summary>
    /// <remarks>
    /// The built-in ids count as headings until the file defines them otherwise: a document a program generated names the
    /// style <c>Heading1</c> and leaves the styles part out, and Word itself shows such a paragraph as a heading. A style the
    /// file does define is a heading only if its outline level or its name says so.
    /// </remarks>
    private static IReadOnlyDictionary<string, int> HeadingLevels(StyleDefinitionsPart? styles)
    {
        Dictionary<string, int> levels = new(BuiltInHeadings, StringComparer.Ordinal);

        if (styles is null)
            return levels;

        using var reader = OpenXmlReader.Create(styles);

        while (reader.Read())
        {
            if (!reader.IsStartElement || reader.ElementType != typeof(W.Style))
                continue;

            var style = (W.Style)reader.LoadCurrentElement()!;

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

    /// <summary>The list and level pairs whose numbering draws a bullet rather than a number, one definition built at a time.</summary>
    private static IReadOnlySet<(int, int)> BulletLevels(NumberingDefinitionsPart? numbering)
    {
        HashSet<(int, int)> bullets = [];

        if (numbering is null)
            return bullets;

        Dictionary<int, HashSet<int>> bulletLevelsOfAbstract = [];
        List<(int List, int Abstract)> instances = [];
        using var reader = OpenXmlReader.Create(numbering);

        while (reader.Read())
        {
            if (!reader.IsStartElement)
                continue;

            if (reader.ElementType == typeof(W.AbstractNum) && reader.LoadCurrentElement() is W.AbstractNum definition && definition.AbstractNumberId?.Value is { } id)
                bulletLevelsOfAbstract[id] = BulletLevels(definition);
            else if (reader.ElementType == typeof(W.NumberingInstance) && reader.LoadCurrentElement() is W.NumberingInstance instance && instance.NumberID?.Value is { } list && instance.AbstractNumId?.Val?.Value is { } abstractId)
                instances.Add((list, abstractId));
        }

        foreach (var (list, abstractId) in instances)
        {
            if (!bulletLevelsOfAbstract.TryGetValue(abstractId, out var levels))
                continue;

            foreach (int level in levels)
                bullets.Add((list, level));
        }

        return bullets;
    }

    private static HashSet<int> BulletLevels(W.AbstractNum definition)
    {
        HashSet<int> levels = [];

        foreach (var level in definition.Elements<W.Level>())
        {
            if (level.NumberingFormat?.Val?.Value == W.NumberFormatValues.Bullet)
                levels.Add(level.LevelIndex?.Value ?? 0);
        }

        return levels;
    }

    [GeneratedRegex(@"^heading\s+(?<level>[1-6])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex HeadingName();

    /// <summary>What one document's walk carries: the heading level of each style, the list levels that draw bullets, the budget of characters and the token.</summary>
    private sealed record Shape(IReadOnlyDictionary<string, int> Levels, IReadOnlySet<(int, int)> Bullets, CharacterBudget Budget, CancellationToken Token);
}
