namespace Snail.Toolkit.Documents;

/// <summary>Reads one format of file into the blocks every format shares.</summary>
/// <remarks>
/// The stream is already in memory when a reader is called, so most formats are read synchronously inside the task. The
/// token is what bounds them: a reader checks it between the rows, pages, slides or elements it walks, and HTML is
/// parsed with it, since a hostile page keeps a parser busy for minutes. <see cref="DocumentReader"/> turns each of these
/// failures into a <see cref="Reading"/> that says why.
/// </remarks>
public interface IFormatReader
{
    /// <summary>The format this reader reads.</summary>
    DocumentKind Format { get; }

    /// <summary>The file extensions it takes, with their dots, lower case.</summary>
    IReadOnlyList<string> Extensions { get; }

    /// <summary>The blocks of the document in <paramref name="content"/>.</summary>
    /// <exception cref="InvalidDataException">The file is not in this format, or is damaged past reading.</exception>
    /// <exception cref="System.Xml.XmlException">A part of an Office file is broken XML; the package around it opened, and the part fails when it is first read.</exception>
    /// <exception cref="DocumentLockedException">The file is protected by a password.</exception>
    /// <exception cref="DocumentTooLargeException">The file unpacks to more than one document may come to.</exception>
    /// <exception cref="OperationCanceledException">The token was cancelled before the reading ended.</exception>
    Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default);
}
