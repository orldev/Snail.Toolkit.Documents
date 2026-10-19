namespace Snail.Toolkit.Documents.Tests;

/// <summary>The pictures of the pages a PDF has no words for, ready for a model that sees.</summary>
public sealed class ScannedPagesTests
{
    /// <summary>The second page of the sample is a picture and nothing else: its picture comes back as a PNG a model can be sent.</summary>
    [Fact]
    public async Task Pictures_PageThatIsAPicture_ComesBackAsAPng()
    {
        var reading = await DocumentReader.Every.ReadAsync(new MemoryStream(Samples.Pdf()), "scan.pdf");
        int[] scans = [.. reading.Document!.Parts.Where(part => part.IsScanned).Select(part => part.Number)];

        var page = Assert.Single(ScannedPages.Pictures(new MemoryStream(Samples.Pdf()), scans));

        Assert.Equal(2, page.Number);
        Assert.Equal("image/png", page.Picture!.MediaType);
        Assert.Equal([0x89, 0x50, 0x4E, 0x47], page.Picture.Bytes[..4]);
    }

    /// <summary>A page of text has no picture to send: it comes back named, without one.</summary>
    [Fact]
    public void Pictures_PageWithoutAnImage_ComesBackWithoutAPicture() =>
        Assert.Null(Assert.Single(ScannedPages.Pictures(new MemoryStream(Samples.Pdf()), [1])).Picture);

    [Fact]
    public void Pictures_PdfUnderAPassword_IsRefusedAsLocked() =>
        Assert.Throws<DocumentLockedException>(() => ScannedPages.Pictures(new MemoryStream(Samples.LockedPdf()), [1]));

    [Fact]
    public void Pictures_CancelledBeforeTheFirstPage_Throws() =>
        Assert.Throws<OperationCanceledException>(() => ScannedPages.Pictures(new MemoryStream(Samples.Pdf()), [2], new CancellationToken(canceled: true)));
}
