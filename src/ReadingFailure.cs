namespace Snail.Toolkit.Documents;

/// <summary>Why a file was not read: each a different answer for whoever handed it in.</summary>
public enum ReadingFailure
{
    /// <summary>No reader takes a file of this name: the format is not one this library reads.</summary>
    UnknownFormat,

    /// <summary>The file is damaged, or is not what its name says.</summary>
    Damaged,

    /// <summary>The file is protected by a password, and reading it takes the password.</summary>
    Locked,

    /// <summary>The file unpacks to more than <see cref="ReadingLimits"/> allow.</summary>
    TooLarge,

    /// <summary>Reading took longer than <see cref="ReadingLimits.MostReadingSeconds"/>.</summary>
    TooSlow
}
