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

    /// <summary>
    /// A PDF may name a picture of any size, and decoding it costs several bytes a pixel: a page of fifty-six million pixels comes back
    /// without a picture for the price of parsing the page, not of decoding 56 MB. The picture is measured on this thread so that only
    /// its decoding is counted, and after a small PDF is read: the first PDF a process opens builds PdfPig's tables, 30 MB once, and a
    /// run that put this test first counted them as the page's.
    /// </summary>
    [Fact]
    public void Pictures_PageWhosePictureIsPastThePixelBudget_ComesBackWithoutAPictureCheaply()
    {
        byte[] pdf = Samples.PdfOfOneHugePicture(8_000, 7_000);
        _ = ScannedPages.Pictures(new MemoryStream(Samples.Pdf()), [1]);
        long before = GC.GetAllocatedBytesForCurrentThread();

        var page = Assert.Single(ScannedPages.Pictures(new MemoryStream(pdf), [1]));

        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.Null(page.Picture);
        Assert.True(allocated < 8 * 1024 * 1024, $"a page of 56 megapixels cost {allocated / 1_048_576} MB without a picture to show for it");
    }

    [Fact]
    public void Pictures_CancelledBeforeTheFirstPage_Throws() =>
        Assert.Throws<OperationCanceledException>(() => ScannedPages.Pictures(new MemoryStream(Samples.Pdf()), [2], new CancellationToken(canceled: true)));
}
