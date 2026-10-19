using DocumentFormat.OpenXml.Packaging;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>Opens an Office file read-only, and turns every way it can fail to open into one failure the server reports.</summary>
/// <remarks>
/// The SDK says a file is not a package in several ways — a damaged zip, a missing part, a zip that is some
/// other format — and none of them is a fault of this server: each means the file is not what its name says.
/// </remarks>
internal static class OfficePackage
{
    public static T Opened<T>(Func<T> open)
    {
        try
        {
            return open();
        }
        catch (Exception failure) when (failure is OpenXmlPackageException or FileFormatException or InvalidDataException or InvalidOperationException)
        {
            throw new InvalidDataException(failure.Message, failure);
        }
    }
}
