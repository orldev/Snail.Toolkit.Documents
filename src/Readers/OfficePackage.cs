using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>Opens an Office file read-only once it is known to unpack within the limits, and turns every way it can fail to open into one failure.</summary>
/// <remarks>
/// The SDK says a file is not a package in several ways — a damaged zip, a missing part, a zip that is some other format —
/// and none of them is a fault of the reader: each means the file is not what its name says.
/// </remarks>
internal static class OfficePackage
{
    /// <summary>The package in <paramref name="content"/>, opened by <paramref name="open"/> after its size unpacked is checked.</summary>
    /// <exception cref="DocumentTooLargeException">The package unpacks past <see cref="ReadingLimits.MostUnpackedBytes"/> or <see cref="ReadingLimits.MostUnpackingRatio"/>.</exception>
    /// <exception cref="InvalidDataException">The file is not a package, or is damaged past opening.</exception>
    public static T Opened<T>(Stream content, Func<Stream, T> open)
    {
        var package = Seekable(content);
        Measure(package);

        try
        {
            return open(package);
        }
        catch (Exception failure) when (failure is OpenXmlPackageException or FileFormatException or InvalidDataException or InvalidOperationException)
        {
            throw new InvalidDataException(failure.Message, failure);
        }
    }

    /// <summary>Refuses a package whose parts unpack past the limits, before the SDK unpacks any of them.</summary>
    /// <remarks>
    /// An Office file is a zip, and a zip of a few kilobytes can unpack to gigabytes, which the SDK would try. The sizes the
    /// zip's directory declares are read first and refuse a truthful bomb for nothing; but a directory is a claim, and a bomb
    /// that claimed a thousand bytes for sixty-four megabytes of zeros passed, and the SDK buffered 47 MB of it before the
    /// zeros failed as XML. So every entry is then inflated into nothing and counted, and the count is refused the moment it
    /// passes a limit; an entry that holds more than its directory declares is refused by the zip reader itself within the
    /// first chunk read. Either way the refusal costs one chunk, and a real document one extra inflation, which the Word and
    /// Excel benchmarks did not register. A file that is not a zip at all — an encrypted Office file is a compound document —
    /// is left for the SDK to name.
    /// </remarks>
    private static void Measure(Stream package)
    {
        if (!IsZip(package))
            return;

        long packed = package.Length;

        using (var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
        {
            Refuse(archive.Entries.Sum(entry => entry.Length), packed);

            long unpacked = 0;

            foreach (var entry in archive.Entries)
                unpacked = Unpacked(entry, unpacked, packed);
        }

        package.Position = 0;
    }

    /// <summary>The bytes unpacked so far after this entry, each piece of it refused as soon as the total passes a limit.</summary>
    private static long Unpacked(ZipArchiveEntry entry, long before, long packed)
    {
        Span<byte> sink = stackalloc byte[16 * 1024];
        long total = before;
        using var inflated = entry.Open();

        for (int read = inflated.Read(sink); read > 0; read = inflated.Read(sink))
        {
            total += read;
            Refuse(total, packed);
        }

        return total;
    }

    private static void Refuse(long unpacked, long packed)
    {
        if (unpacked > ReadingLimits.MostUnpackedBytes || unpacked > packed * ReadingLimits.MostUnpackingRatio)
            throw new DocumentTooLargeException($"the file of {packed} bytes would unpack to {unpacked}");
    }

    private static bool IsZip(Stream package)
    {
        Span<byte> signature = stackalloc byte[4];
        int read = package.ReadAtLeast(signature, signature.Length, throwOnEndOfStream: false);
        package.Position = 0;

        return read == 4 && signature is [0x50, 0x4B, 0x03, 0x04];
    }

    private static Stream Seekable(Stream content)
    {
        if (content.CanSeek)
            return content;

        var copy = new MemoryStream();
        content.CopyTo(copy);
        copy.Position = 0;

        return copy;
    }
}
