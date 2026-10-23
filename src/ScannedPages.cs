using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Snail.Toolkit.Documents.Readers;

namespace Snail.Toolkit.Documents;

/// <summary>A picture a model that sees can be sent: the bytes of an image and what they are.</summary>
/// <param name="Bytes">The image.</param>
/// <param name="MediaType"><c>image/png</c> or <c>image/jpeg</c>.</param>
public sealed record PagePicture(byte[] Bytes, string MediaType);

/// <summary>A page of a PDF without a text layer, and its picture; no picture when no decoder opens it.</summary>
public sealed record ScannedPage(int Number, PagePicture? Picture);

/// <summary>The pictures of the scanned pages of a PDF, for a model that reads what a scan shows.</summary>
/// <remarks>
/// A scanned page has no words in the document's text — <see cref="Part.IsScanned"/> says which — and its picture is
/// the only way to them. A scan is one full-page picture, so the largest image of a page is the page.
/// </remarks>
public static class ScannedPages
{
    /// <summary>The picture of each named page of the PDF, in a form a model can be sent; null where it cannot be decoded.</summary>
    /// <remarks>
    /// Fax and JBIG2 encodings, common in black-and-white scans, have no managed decoder, and such a page comes back
    /// without a picture rather than as bytes a model cannot look at; so does a picture past
    /// <see cref="ReadingLimits.MostPicturePixels"/>, which would be decoded into more memory than a page may take. A
    /// decoder throws on input it does not expect, and one broken image costs its own page, never the pictures of the others.
    /// </remarks>
    /// <exception cref="InvalidDataException">The file is not a PDF, or is damaged past opening.</exception>
    /// <exception cref="DocumentLockedException">The PDF is protected by a password.</exception>
    /// <exception cref="OperationCanceledException">The token was cancelled before every page was decoded.</exception>
    public static IReadOnlyList<ScannedPage> Pictures(Stream pdf, IReadOnlyList<int> pages, CancellationToken cancellationToken = default)
    {
        using var document = PdfReader.Opened(pdf);
        List<ScannedPage> scanned = [];

        foreach (int number in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            scanned.Add(new ScannedPage(number, Decoded(document, number)));
        }

        return scanned;
    }

    private static PagePicture? Decoded(PdfDocument document, int number)
    {
        try
        {
            return Largest(document.GetPage(number).GetImages());
        }
        catch (Exception failure) when (failure is not OutOfMemoryException)
        {
            return null;
        }
    }

    private static PagePicture? Largest(IEnumerable<IPdfImage> images)
    {
        var image = images.MaxBy(candidate => candidate.BoundingBox.Width * candidate.BoundingBox.Height);

        if (image is null || (long)image.WidthInSamples * image.HeightInSamples > ReadingLimits.MostPicturePixels)
            return null;

        if (image.TryGetPng(out var png))
            return new PagePicture(png, "image/png");

        byte[] raw = [.. image.RawBytes];

        return raw is [0xFF, 0xD8, ..] ? new PagePicture(raw, "image/jpeg") : null;
    }
}
