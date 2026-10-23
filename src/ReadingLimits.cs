namespace Snail.Toolkit.Documents;

/// <summary>How far one document is read: what a file may unpack to and how long it may take, whatever its size on disk.</summary>
/// <remarks>
/// A file past any of these is refused with <see cref="DocumentTooLargeException"/>, never cut short in silence: a reader
/// handed the first part of a document as if it were the whole has been told something false about it.
/// </remarks>
public static class ReadingLimits
{
    /// <summary>Characters the whole text of one document may come to; a file that unpacks to more is refused, not read.</summary>
    /// <remarks>
    /// The size of a file says nothing about the size of its text: a workbook of 11 648 bytes with one cell at XFD on
    /// each row came to 98 408 325 characters and 4.69 GB allocated. Twenty million characters are forty megabytes of
    /// text, which a container of 512 MB holds beside the documents a server keeps in memory.
    /// </remarks>
    public const int MostDocumentCharacters = 20_000_000;

    /// <summary>Cells one table may come to, counting the empty ones a short row is padded with.</summary>
    /// <remarks>
    /// Padding is where a small file grows: rows that each reach the last column of a sheet, or records that each name
    /// a key of their own, become grids of millions of empty cells. A million cells is a sheet of fifty thousand rows
    /// of twenty columns.
    /// </remarks>
    public const int MostCells = 1_000_000;

    /// <summary>Seconds one file may take to read before the reading is stopped and refused.</summary>
    /// <remarks>
    /// The budget of characters bounds what a file becomes, not how long it takes to become it: a sparse workbook took 27 s
    /// before the budget refused it, and a page of fifty thousand nested elements keeps the HTML parser busy for some
    /// ninety. Half a minute is several times the slowest real document of the benchmarks, a PDF of twenty pages in 17 ms.
    /// </remarks>
    public const int MostReadingSeconds = 30;

    /// <summary>Bytes an Office file may unpack to, all its parts together.</summary>
    /// <remarks>Word, Excel and PowerPoint files are zip archives: a few kilobytes can unpack to gigabytes, and a reader would try.</remarks>
    public const long MostUnpackedBytes = 256L * 1024 * 1024;

    /// <summary>How many times its own size an Office file may unpack to.</summary>
    /// <remarks>
    /// The limit the worker of Snail.Swarm.Agent refused Office files with before it read them through this library, kept
    /// as it was together with <see cref="MostUnpackedBytes"/>: a document's XML compresses well, but a bomb compresses by
    /// thousands — eight megabytes of zeros pack into some eight kilobytes.
    /// </remarks>
    public const int MostUnpackingRatio = 100;

    /// <summary>Pixels a scanned page's picture may have to be decoded; a larger one comes back without a picture.</summary>
    /// <remarks>
    /// A page is decoded to a PNG in memory, several bytes a pixel: an A4 page scanned at 600 dpi is 35 million pixels and some
    /// 140 MB decoded, at 1200 dpi four times that, and a PDF may name a picture of any size at all. Fifty million pixels take
    /// every scan a person makes and leave the one made to exhaust memory undecoded.
    /// </remarks>
    public const long MostPicturePixels = 50_000_000;

    /// <summary>Levels a list item is indented at most; Word and PowerPoint number nine, from 0 to 8.</summary>
    /// <remarks>A level is an integer in the file, and indenting by it would turn one item into a line of millions of spaces.</remarks>
    public const int MostListDepth = 8;
}
