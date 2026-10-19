namespace Snail.Toolkit.Documents;

/// <summary>A file small enough on disk whose content, once read, comes to more than one document may.</summary>
/// <remarks>
/// Told apart from a damaged file because the answer differs: a damaged file is asked for again, a document this large
/// is split or read as a smaller export. It is the second failure a reader may throw besides
/// <see cref="InvalidDataException"/>, as <see cref="IFormatReader"/> says.
/// </remarks>
public sealed class DocumentTooLargeException(string message) : Exception(message);
