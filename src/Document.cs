namespace Snail.Toolkit.Documents;

/// <summary>A document as this server serves it: its text as Markdown and where everything in it stands.</summary>
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
    /// <summary>The widest row: the number of columns the table is read with.</summary>
    public int Columns => Rows.Count == 0 ? 0 : Rows.Max(row => row.Count);
}
