namespace Snail.Toolkit.Documents.Readers;

/// <summary>Counts the characters a reader has produced and refuses the file at the first block past what one document may hold.</summary>
/// <remarks>
/// The writer refuses a document past <see cref="ReadingLimits.MostDocumentCharacters"/> too, but only once every block stands in
/// memory; counted here, a reader stops at the budget with the blocks it has, so a PDF of a thousand dense pages costs the budget,
/// not the file. Text-based formats need no counter: their bytes are bounded before they are decoded.
/// </remarks>
internal sealed class CharacterBudget
{
    private long _counted;

    /// <summary>Counts a block's text; the file is refused when the total passes the budget.</summary>
    /// <exception cref="DocumentTooLargeException">The text read so far comes to more than one document may.</exception>
    public void Count(string text)
    {
        _counted += text.Length;

        if (_counted > ReadingLimits.MostDocumentCharacters)
            throw new DocumentTooLargeException($"its text comes to more than {ReadingLimits.MostDocumentCharacters} characters");
    }
}
