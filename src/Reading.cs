using System.Diagnostics.CodeAnalysis;

namespace Snail.Toolkit.Documents;

/// <summary>What came of reading one file: the document, or why it was not read and what the reader said.</summary>
/// <remarks>
/// Made only by <see cref="Of"/> and <see cref="Refused"/>, so a reading holds a document or a failure with its reason,
/// never both and never neither; <see cref="IsRead"/> tells the compiler which.
/// </remarks>
public sealed record Reading
{
    private Reading(Document? document, ReadingFailure? failure, string? reason)
    {
        Document = document;
        Failure = failure;
        Reason = reason;
    }

    /// <summary>The document; null when it was not read.</summary>
    public Document? Document { get; }

    /// <summary>Why it was not read; null when it was.</summary>
    public ReadingFailure? Failure { get; }

    /// <summary>What the reader said about the failure, in its own words; null when it was read.</summary>
    public string? Reason { get; }

    /// <summary>Whether the file was read: the document is there when it was, the failure and its reason when it was not.</summary>
    [MemberNotNullWhen(true, nameof(Document))]
    [MemberNotNullWhen(false, nameof(Failure), nameof(Reason))]
    public bool IsRead => Document is not null;

    /// <summary>The document of a file that was read.</summary>
    /// <remarks>
    /// A false answer states that <see cref="Failure"/> and <see cref="Reason"/> are there, so the branch that refuses the
    /// file reads them without a null check — one call answers both ways, as <c>Outcome.TryGetValue</c> does in Snail.Toolkit.Web.
    /// </remarks>
    [MemberNotNullWhen(false, nameof(Failure), nameof(Reason))]
    public bool TryGetDocument([NotNullWhen(true)] out Document? document)
    {
        document = Document;

        return document is not null;
    }

    /// <summary>A file read into its document.</summary>
    public static Reading Of(Document document) => new(document, null, null);

    /// <summary>A file not read, and why.</summary>
    public static Reading Refused(ReadingFailure failure, string reason) => new(null, failure, reason);
}
