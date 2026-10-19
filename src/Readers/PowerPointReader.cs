using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>A PowerPoint presentation: every slide a part titled by its title, then its text, tables and notes.</summary>
/// <remarks>
/// The speaker's notes are read too and marked as notes: they often hold the half of the talk the slide leaves
/// out, and a model asked what the presentation says would otherwise answer from the bullets alone.
/// </remarks>
public sealed class PowerPointReader : IFormatReader
{
    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.PowerPoint;

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
        int number = 0;

        foreach (var id in presentation.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            number++;

            if (id.RelationshipId?.Value is not { } relationship || presentation.GetPartById(relationship) is not SlidePart slide)
                continue;

            var shapes = slide.Slide?.CommonSlideData?.ShapeTree;
            var title = shapes?.Descendants<P.Shape>().FirstOrDefault(IsTitle);

            blocks.Add(new PartStart(PartKind.Slide, number, title is null ? null : Text(title.TextBody)));

            if (shapes is not null)
                blocks.AddRange(Blocks(shapes));

            if (slide.NotesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree is { } notes
                && string.Join('\n', notes.Descendants<P.Shape>().Where(IsBody).Select(shape => Text(shape.TextBody))) is { Length: > 0 } said)
                blocks.Add(new Paragraph($"Notes: {said}"));
        }

        return blocks;
    }

    private static IEnumerable<Block> Blocks(OpenXmlElement container)
    {
        foreach (var element in container.Elements())
        {
            switch (element)
            {
                case P.Shape shape when !IsTitle(shape) && shape.TextBody is { } body:
                    foreach (var paragraph in body.Elements<A.Paragraph>())
                    {
                        if (Text(paragraph) is not { Length: > 0 } text)
                            continue;

                        yield return IsBody(shape)
                            ? new ListItem(paragraph.ParagraphProperties?.Level?.Value ?? 0, text, IsOrdered: false)
                            : new Paragraph(text);
                    }

                    break;
                case P.GraphicFrame frame when frame.Descendants<A.Table>().FirstOrDefault() is { } table:
                    yield return new TableBlock(null, Rows(table));
                    break;
                case P.GroupShape group:
                    foreach (var inner in Blocks(group))
                        yield return inner;
                    break;
            }
        }
    }

    private static IReadOnlyList<IReadOnlyList<string>> Rows(A.Table table) =>
    [
        .. table.Elements<A.TableRow>()
            .Select(row => (IReadOnlyList<string>)[.. row.Elements<A.TableCell>().Select(cell => Text(cell.TextBody))])
    ];

    private static bool IsTitle(P.Shape shape) =>
        Placeholder(shape) is { } type && (type == P.PlaceholderValues.Title || type == P.PlaceholderValues.CenteredTitle);

    /// <summary>A placeholder of the layout's body, whose lines are the slide's points; a placeholder with no type is one.</summary>
    private static bool IsBody(P.Shape shape) =>
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is { } placeholder
        && (placeholder.Type?.Value is null || placeholder.Type.Value == P.PlaceholderValues.Body || placeholder.Type.Value == P.PlaceholderValues.Object);

    private static P.PlaceholderValues? Placeholder(P.Shape shape) =>
        shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;

    private static string Text(OpenXmlElement? element) =>
        element is null
            ? string.Empty
            : string.Join('\n', element.Descendants<A.Paragraph>().Select(Text).Where(line => line.Length > 0));

    private static string Text(A.Paragraph paragraph) =>
        string.Concat(paragraph.Descendants().Select(piece => piece switch
        {
            A.Text text => text.Text,
            A.Break => "\n",
            _ => string.Empty
        })).Trim();
}
