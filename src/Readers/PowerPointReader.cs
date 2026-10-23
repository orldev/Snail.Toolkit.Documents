using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>A PowerPoint presentation: every slide a part titled by its title, then its text, tables and notes.</summary>
/// <remarks>
/// The speaker's notes are read too and marked as notes: they often hold the half of the talk the slide leaves
/// out, and a model asked what the presentation says would otherwise answer from the bullets alone.
/// Nothing of a slide is built as a tree but a placeholder's type and a paragraph's properties: shapes, paragraphs
/// and tables are read as the part reader's events, their text gathered piece by piece under the budget of
/// characters, so what the heap holds is the budget — one shape of a hundred million letters stood at 339 MB as a tree.
/// </remarks>
public sealed class PowerPointReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Kind => DocumentKind.PowerPoint;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".pptx", ".pptm"];

    /// <inheritdoc/>
    public Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default) =>
        Task.FromResult(Read(content, cancellationToken));

    private static IReadOnlyList<Block> Read(Stream content, CancellationToken cancellationToken)
    {
        using var deck = OfficePackage.Opened(content, package => PresentationDocument.Open(package, isEditable: false));

        var presentation = deck.PresentationPart ?? throw new InvalidDataException("the file holds no presentation");
        List<Block> blocks = [];
        var budget = new CharacterBudget();
        int number = 0;

        foreach (var id in Slides(presentation, cancellationToken))
        {
            number++;

            if (id.RelationshipId?.Value is not { } relationship)
                continue;

            if (!presentation.TryGetPartById(relationship, out var part))
                throw new InvalidDataException($"the presentation lists slide {number} under a relationship it does not hold");

            if (part is SlidePart slide)
                blocks.AddRange(Blocks(slide, number, budget, cancellationToken));
        }

        return blocks;
    }

    /// <summary>The slides the presentation lists, in order, one entry of the list built at a time.</summary>
    /// <remarks>
    /// The list is the presentation's table of contents, a few lines in a real file; two million entries, sixty-five megabytes the zip
    /// limits let through, stood at 779 MB as a tree before the first slide was looked up. Read entry by entry, every slide's heading is
    /// counted against the budget of characters as it is met.
    /// </remarks>
    private static IEnumerable<P.SlideId> Slides(PresentationPart presentation, CancellationToken cancellationToken)
    {
        using var reader = OpenXmlReader.Create(presentation);

        while (reader.Read())
        {
            if (!reader.IsStartElement || reader.ElementType != typeof(P.SlideId))
                continue;

            cancellationToken.ThrowIfCancellationRequested();

            yield return (P.SlideId)reader.LoadCurrentElement()!;
        }
    }

    /// <summary>The blocks of one slide: its start under its title, its points and tables in order, then the speaker's notes.</summary>
    private static IEnumerable<Block> Blocks(SlidePart slide, int number, CharacterBudget budget, CancellationToken cancellationToken)
    {
        string? title = null;
        List<Block> content = [];

        foreach (var figure in Figures(slide, budget, cancellationToken))
        {
            if (figure.IsTitle)
                title ??= Text(figure);
            else if (figure.Rows is { } rows)
                content.Add(new TableBlock(null, rows));
            else
                content.AddRange(Points(figure));
        }

        var start = new PartStart(PartKind.Slide, number, title);
        budget.Count(Markdown.PartLine(start));

        yield return start;

        foreach (var block in content)
            yield return block;

        if (Notes(slide.NotesSlidePart, budget, cancellationToken) is { Length: > 0 } notes)
            yield return new Paragraph($"Notes: {notes}");
    }

    /// <summary>The lines of a figure as blocks: points of the slide when the figure is the layout's body, paragraphs otherwise.</summary>
    private static IEnumerable<Block> Points(Figure figure)
    {
        foreach (var (level, text) in figure.Lines)
        {
            yield return figure.IsBody
                ? new ListItem(level, text, IsOrdered: false)
                : new Paragraph(text);
        }
    }

    private static string Notes(NotesSlidePart? notes, CharacterBudget budget, CancellationToken cancellationToken)
    {
        if (notes is null)
            return string.Empty;

        var spoken = Figures(notes, budget, cancellationToken).Where(figure => figure.IsBody).Select(Text);

        return string.Join('\n', spoken);
    }

    /// <summary>The text of a figure, a line a paragraph.</summary>
    private static string Text(Figure figure) =>
        string.Join('\n', figure.Lines.Select(line => line.Text));

    /// <summary>The shapes and tables of a part in order, each read to its end as the reader meets it; shapes in a group arrive as their own.</summary>
    private static IEnumerable<Figure> Figures(OpenXmlPart part, CharacterBudget budget, CancellationToken cancellationToken)
    {
        using var reader = OpenXmlReader.Create(part);

        while (reader.Read())
        {
            if (!reader.IsStartElement)
                continue;

            cancellationToken.ThrowIfCancellationRequested();

            if (reader.ElementType == typeof(P.Shape))
                yield return Shaped(reader, budget);
            else if (reader.ElementType == typeof(A.Table))
                yield return new Figure(IsTitle: false, IsBody: false, [], Rows(reader, budget));
        }
    }

    /// <summary>The shape the reader stands on, read to its end: what placeholder of the layout it is, and its paragraphs' level and text.</summary>
    /// <remarks>A placeholder of the layout's body holds the slide's points; a placeholder with no type is one too.</remarks>
    private static Figure Shaped(OpenXmlReader reader, CharacterBudget budget)
    {
        int depth = reader.Depth;
        bool isPlaceholder = false;
        P.PlaceholderValues? type = null;
        List<(int Level, string Text)> lines = [];

        while (reader.Read() && !reader.IsEndAt(depth))
        {
            if (!reader.IsStartElement)
                continue;

            if (reader.ElementType == typeof(P.PlaceholderShape) && reader.LoadCurrentElement() is P.PlaceholderShape placeholder)
            {
                isPlaceholder = true;
                type = placeholder.Type?.Value;
            }
            else if (reader.ElementType == typeof(A.Paragraph) && Line(reader, budget) is { Text.Length: > 0 } line)
            {
                lines.Add(line);
            }
        }

        bool isTitle = type is { } placed && (placed == P.PlaceholderValues.Title || placed == P.PlaceholderValues.CenteredTitle);
        bool isBody = isPlaceholder && (type is null || type == P.PlaceholderValues.Body || type == P.PlaceholderValues.Object);

        return new Figure(isTitle, isBody, lines, null);
    }

    /// <summary>The paragraph the reader stands on, read to its end: its level and its text, pieces and line breaks gathered under the budget.</summary>
    private static (int Level, string Text) Line(OpenXmlReader reader, CharacterBudget budget)
    {
        int depth = reader.Depth;
        int level = 0;
        var text = new StringBuilder();

        while (reader.Read() && !reader.IsEndAt(depth))
        {
            if (!reader.IsStartElement)
                continue;

            if (reader.ElementType == typeof(A.ParagraphProperties) && reader.Depth == depth + 1 && reader.LoadCurrentElement() is A.ParagraphProperties properties)
            {
                level = properties.Level?.Value ?? 0;
            }
            else if (reader.ElementType == typeof(A.Text))
            {
                string piece = reader.GetText();
                budget.Count(piece);
                text.Append(piece);
            }
            else if (reader.ElementType == typeof(A.Break))
            {
                text.Append('\n');
            }
        }

        return (level, text.ToString().Trim());
    }

    /// <summary>The rows of the table the reader stands on, read to its end, each cell its lines.</summary>
    private static IReadOnlyList<IReadOnlyList<string>> Rows(OpenXmlReader reader, CharacterBudget budget)
    {
        int depth = reader.Depth;
        List<IReadOnlyList<string>> rows = [];
        List<string> cells = [];

        while (reader.Read() && !reader.IsEndAt(depth))
        {
            if (reader.IsStartElement && reader.ElementType == typeof(A.TableRow) && reader.Depth == depth + 1)
                cells = [];
            else if (reader.IsStartElement && reader.ElementType == typeof(A.TableCell) && reader.Depth == depth + 2)
                cells.Add(CellText(reader, budget));
            else if (reader.IsEndElement && reader.ElementType == typeof(A.TableRow) && reader.Depth == depth + 1)
                rows.Add(cells);
        }

        return rows;
    }

    private static string CellText(OpenXmlReader reader, CharacterBudget budget)
    {
        int depth = reader.Depth;
        List<string> lines = [];

        while (reader.Read() && !reader.IsEndAt(depth))
        {
            if (reader.IsStartElement && reader.ElementType == typeof(A.Paragraph) && Line(reader, budget) is { Text.Length: > 0 } line)
                lines.Add(line.Text);
        }

        return string.Join('\n', lines);
    }

    /// <summary>A shape or a table of a slide as read: whether it is the layout's title or body, its lines, or its rows when it is a table.</summary>
    private sealed record Figure(bool IsTitle, bool IsBody, IReadOnlyList<(int Level, string Text)> Lines, IReadOnlyList<IReadOnlyList<string>>? Rows);
}
