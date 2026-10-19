namespace Snail.Toolkit.Documents;

/// <summary>What a document is divided into, when it is divided at all.</summary>
public enum PartKind
{
    /// <summary>A page of a PDF.</summary>
    Page,

    /// <summary>A sheet of a workbook.</summary>
    Sheet,

    /// <summary>A slide of a presentation.</summary>
    Slide
}
