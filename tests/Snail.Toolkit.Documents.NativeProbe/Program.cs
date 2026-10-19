using System.Text;
using Snail.Toolkit.Documents;
using Snail.Toolkit.Documents.Readers;

var table = await new DelimitedReader().ReadAsync(new MemoryStream(Encoding.UTF8.GetBytes("name;sum\nAlpha;1042\n")));
var notes = await new MarkdownReader().ReadAsync(new MemoryStream(Encoding.UTF8.GetBytes("# Report\n\nPaid in full.\n")));
var document = Markdown.Of(DocumentKind.Csv, [.. table, .. notes]);

Console.WriteLine(document.Text);

return document.Tables.Count == 1 && document.Headings.Count == 1 && document.Text.Contains("1042", StringComparison.Ordinal) ? 0 : 1;
