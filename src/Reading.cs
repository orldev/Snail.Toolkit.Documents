namespace Snail.Toolkit.Documents;

/// <summary>What came of reading one file: the document, or why it was not read and what the reader said.</summary>
/// <param name="Document">The document; null when it was not read.</param>
/// <param name="Failure">Why it was not read; null when it was.</param>
/// <param name="Reason">What the reader said about the failure, in its own words; empty when it was read.</param>
public sealed record Reading(Document? Document, ReadingFailure? Failure, string Reason)
{
    /// <summary>Whether the file was read.</summary>
    public bool IsRead => Document is not null;

    /// <summary>A file read into its document.</summary>
    public static Reading Of(Document document) => new(document, null, string.Empty);

    /// <summary>A file not read, and why.</summary>
    public static Reading Refused(ReadingFailure failure, string reason) => new(null, failure, reason);
}
