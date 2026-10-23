using System.Text;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>The text of a file whose encoding nobody wrote down.</summary>
/// <remarks>
/// A byte-order mark is believed. Without one the bytes are tried as UTF-8 strictly, and only when they are not
/// valid UTF-8 are they read as Windows-1251 — the encoding Excel and Notepad on a Russian Windows save a CSV in,
/// which as UTF-8 turns every Cyrillic letter into a replacement character and the table into noise.
/// </remarks>
internal static class Decoding
{
    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    static Decoding() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <summary>The text of the stream from where it stands, in the encoding its bytes say.</summary>
    /// <remarks>
    /// UTF-8 and Windows-1251 give at least one character per byte and UTF-16 one per two, so a file of more than twice the
    /// budget of characters in bytes cannot fit whatever it is encoded in and is refused before a byte is decoded: decoding
    /// a 50 MB text file to refuse it cost 809 MB. A file between the budget and twice it is decoded and refused by the writer.
    /// </remarks>
    /// <exception cref="DocumentTooLargeException">The stream holds more than twice <see cref="ReadingLimits.MostDocumentCharacters"/> bytes.</exception>
    public static string Text(Stream content)
    {
        if (content.CanSeek && content.Length - content.Position > 2L * ReadingLimits.MostDocumentCharacters)
            throw new DocumentTooLargeException($"a text file of {content.Length - content.Position} bytes holds more than {ReadingLimits.MostDocumentCharacters} characters");

        var bytes = Bytes(content);

        if (Marked(bytes) is { } marked)
            return marked;

        try
        {
            return StrictUtf8.GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return Encoding.GetEncoding(1251).GetString(bytes);
        }
    }

    /// <summary>The bytes of the stream from where it stands, without a copy when they are already in memory.</summary>
    private static ArraySegment<byte> Bytes(Stream content)
    {
        if (content is MemoryStream memory && memory.TryGetBuffer(out var buffer))
            return buffer[(int)memory.Position..];

        using var copy = new MemoryStream();
        content.CopyTo(copy);

        return copy.ToArray();
    }

    private static string? Marked(ReadOnlySpan<byte> bytes)
    {
        if (bytes is [0xEF, 0xBB, 0xBF, ..])
            return Encoding.UTF8.GetString(bytes[3..]);

        if (bytes is [0xFF, 0xFE, ..])
            return Encoding.Unicode.GetString(bytes[2..]);

        return bytes is [0xFE, 0xFF, ..] ? Encoding.BigEndianUnicode.GetString(bytes[2..]) : null;
    }
}
