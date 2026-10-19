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
    /// zip declares are what its entries are read up to, so they are read from its directory and nothing is unpacked. A
    /// file that is not a zip at all — an encrypted Office file is a compound document — is left for the SDK to name.
    /// </remarks>
    private static void Measure(Stream package)
    {
        if (!IsZip(package))
            return;

        long packed = package.Length;
        long unpacked;

        using (var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
            unpacked = archive.Entries.Sum(entry => entry.Length);

        package.Position = 0;

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
