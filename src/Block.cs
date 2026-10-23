namespace Snail.Toolkit.Documents;

/// <summary>One piece of a document's structure, as a format adapter reads it and before it is written as Markdown.</summary>
/// <remarks>
/// The adapters say what a document holds and never how it is written: every format turns into the same few
/// blocks, so headings, tables and the marks of pages look alike whatever the file was, and the outline and the
/// tables are read from the same pass that writes the text.
/// </remarks>
public abstract record Block
{
    /// <summary>Only the blocks below are blocks: a reader of another library builds documents of them, and a kind the writer does not know would be left out of the text.</summary>
    private protected Block()
    {
    }
}

/// <summary>The start of a page, sheet or slide.</summary>
/// <remarks>A scanned page carries pictures and no text layer: its words are not in the document's text.</remarks>
public sealed record PartStart(PartKind Kind, int Number, string? Title, bool IsScanned = false) : Block;

/// <summary>A heading of the document, level 1 to 6.</summary>
public sealed record Heading(int Level, string Text) : Block;

/// <summary>A paragraph of running text.</summary>
public sealed record Paragraph(string Text) : Block;

/// <summary>An item of a list, nested <paramref name="Depth"/> levels deep.</summary>
public sealed record ListItem(int Depth, string Text, bool IsOrdered) : Block;

/// <summary>A table, its first row the header as the document has it.</summary>
public sealed record TableBlock(string? Title, IReadOnlyList<IReadOnlyList<string>> Rows) : Block
{
    /// <summary>Where its cells stand on the sheet; null for a table that is not a sheet's.</summary>
    public CellAddresses? Addresses { get; init; }
}

/// <summary>Text kept exactly as written: code, or data that has no prose shape.</summary>
public sealed record CodeBlock(string Language, string Text) : Block;
