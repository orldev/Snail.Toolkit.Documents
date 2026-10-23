namespace Snail.Toolkit.Documents;

/// <summary>A document read: its text as Markdown and where everything in it stands.</summary>
/// <param name="Kind">What the file was.</param>
/// <param name="Text">The whole text, as Markdown.</param>
/// <param name="Parts">Its pages, sheets or slides, in order; empty for a document that is not divided.</param>
/// <param name="Sections">Its sections, each opened by a heading, in order.</param>
/// <param name="Tables">Its tables, in order.</param>
public sealed record Document(
    DocumentKind Kind,
    string Text,
    IReadOnlyList<Part> Parts,
    IReadOnlyList<Section> Sections,
    IReadOnlyList<Table> Tables)
{
    /// <summary>The document the blocks make: their text written as Markdown, and where each part, heading and table of it starts.</summary>
    /// <remarks>A program with a reader of its own turns that reader's blocks into a document here, the same way <see cref="DocumentReader"/> does.</remarks>
    /// <exception cref="DocumentTooLargeException">The text, or one of its tables, comes to more than one document may.</exception>
    public static Document Of(DocumentKind kind, IReadOnlyList<Block> blocks) =>
        Markdown.Write(kind, blocks);
}

/// <summary>A page, sheet or slide, and where its text starts.</summary>
public sealed record Part(PartKind Kind, int Number, string? Title, int Start, int Length, bool IsScanned);

/// <summary>A section of the text: the heading that opens it, where it starts and how far it runs — to the next heading of its level or above, or to the end.</summary>
/// <remarks>A section holds the sections under it: a level-2 section runs through the level-3 ones that follow, as the outline of a book does.</remarks>
public sealed record Section(int Level, string Title, int Start, int Length);

/// <summary>A table of the document: its number, the part it is on and its rows.</summary>
public sealed record Table(int Number, string? Title, int? Part, int Start, IReadOnlyList<IReadOnlyList<string>> Rows)
{
    /// <summary>Where its cells stand on the sheet; null for a table that is not a sheet's.</summary>
    public CellAddresses? Addresses { get; init; }

    /// <summary>The widest row: the number of columns the table is read with.</summary>
    /// <remarks>Counted once: a table of a million rows is not walked again on every look at its width.</remarks>
    public int Columns { get; } = Rows.Count == 0 ? 0 : Rows.Max(row => row.Count);
}
