namespace Snail.Toolkit.Documents;

/// <summary>A document read: its text as Markdown and where everything in it stands.</summary>
/// <param name="Format">What the file was.</param>
/// <param name="Text">The whole text, as Markdown.</param>
/// <param name="Parts">Its pages, sheets or slides, in order; empty for a document that is not divided.</param>
/// <param name="Headings">Its headings, in order.</param>
/// <param name="Tables">Its tables, in order.</param>
public sealed record Document(
    DocumentKind Format,
    string Text,
    IReadOnlyList<Part> Parts,
    IReadOnlyList<Section> Headings,
    IReadOnlyList<Table> Tables);

/// <summary>A page, sheet or slide, and where its text starts.</summary>
public sealed record Part(PartKind Kind, int Number, string? Title, int Start, int Length, bool IsScanned);

/// <summary>A heading, and where it stands in the text.</summary>
public sealed record Section(int Level, string Title, int Start);

/// <summary>A table of the document: its number, the part it is on and its rows.</summary>
public sealed record Table(int Number, string? Title, int? Part, int Start, IReadOnlyList<IReadOnlyList<string>> Rows)
{
    /// <summary>Where its cells stand on the sheet; null for a table that is not a sheet's.</summary>
    public CellAddresses? Addresses { get; init; }

    /// <summary>The widest row: the number of columns the table is read with.</summary>
    /// <remarks>Counted once: a table of a million rows is not walked again on every look at its width.</remarks>
    public int Columns { get; } = Rows.Count == 0 ? 0 : Rows.Max(row => row.Count);
}
