namespace Snail.Toolkit.Documents;

/// <summary>The file is protected by a password: it is not damaged, and reading it takes the password.</summary>
/// <remarks>
/// Told apart from a damaged file because the answer differs: a damaged file is asked for again, a locked one is asked
/// for without its password.
/// </remarks>
public sealed class DocumentLockedException : Exception
{
    /// <summary>A file refused for its password, with what the reader saw.</summary>
    public DocumentLockedException(string message)
        : base(message)
    {
    }

    /// <summary>A file refused for its password, with the failure of the format's own library.</summary>
    public DocumentLockedException(string message, Exception inner)
        : base(message, inner)
    {
    }
}
