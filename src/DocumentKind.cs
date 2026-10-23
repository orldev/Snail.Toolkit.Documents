namespace Snail.Toolkit.Documents;

/// <summary>The kinds of file this library reads.</summary>
public enum DocumentKind
{
    /// <summary>A Word document, .docx.</summary>
    Word,

    /// <summary>An Excel workbook, .xlsx.</summary>
    Excel,

    /// <summary>A PowerPoint presentation, .pptx.</summary>
    PowerPoint,

    /// <summary>A PDF, read through its text layer.</summary>
    Pdf,

    /// <summary>A web page saved as .html or .htm.</summary>
    Html,

    /// <summary>Values separated by commas, semicolons or tabs: .csv and .tsv.</summary>
    Csv,

    /// <summary>A JSON file.</summary>
    Json,

    /// <summary>A Markdown file.</summary>
    Markdown,

    /// <summary>Plain text.</summary>
    Text
}
