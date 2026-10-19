using System.Text;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>A saved web page: its headings, paragraphs, lists, tables and code, without menus, scripts and styles.</summary>
/// <remarks>
/// Text that sits loose inside a container — a <c>div</c> with words and links but no paragraph — is gathered
/// into a paragraph of its own, because most pages are written that way and a reader that waited for
/// <c>p</c> would find them empty.
/// </remarks>
public sealed class HtmlReader : IFormatReader
{
    private static readonly HashSet<string> Skipped = new(StringComparer.OrdinalIgnoreCase) { "script", "style", "noscript", "template", "svg", "nav", "iframe", "head", "form", "button" };

    private static readonly HashSet<string> Inline = new(StringComparer.OrdinalIgnoreCase) { "a", "span", "strong", "b", "em", "i", "u", "code", "small", "sub", "sup", "mark", "abbr", "time", "label", "br", "img", "s", "q", "cite", "kbd", "var" };

    /// <inheritdoc/>
    public DocumentKind Format => DocumentKind.Html;

    /// <inheritdoc/>
    public IReadOnlyList<string> Extensions { get; } = [".html", ".htm"];

    /// <inheritdoc/>
    /// <remarks>
    /// Parsed with the token: AngleSharp builds a tree of fifty thousand nested elements in some ninety seconds of CPU,
    /// and only its asynchronous parse stops when asked to.
    /// </remarks>
    public async Task<IReadOnlyList<Block>> ReadAsync(Stream content, CancellationToken cancellationToken = default)
    {
        var page = await new HtmlParser().ParseDocumentAsync(Decoding.Text(content), cancellationToken);
        List<Block> blocks = [];

        if (page.Title is { Length: > 0 } title && page.QuerySelector("h1") is null)
            blocks.Add(new Heading(1, title.Trim()));

        Walk(page.Body ?? page.DocumentElement, blocks, cancellationToken);

        return blocks;
    }

    /// <summary>Reads the tree in document order without recursion, so the depth of a page costs heap rather than stack.</summary>
    /// <remarks>
    /// Recursion overflowed the stack of a pool thread at three thousand nested divs, a page of fifteen kilobytes, and a
    /// stack overflow ends the process every session of the machine shares. Loose text is flushed when a container is
    /// entered and when it is left, exactly where the recursive walk flushed it.
    /// </remarks>
    private static void Walk(INode root, List<Block> blocks, CancellationToken cancellationToken)
    {
        var loose = new StringBuilder();
        Stack<IEnumerator<INode>> open = [];
        open.Push(root.ChildNodes.GetEnumerator());

        while (open.TryPeek(out var children))
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!children.MoveNext())
            {
                open.Pop();
                Flush(blocks, loose);
                continue;
            }

            if (children.Current is IText text)
            {
                loose.Append(text.Data);
                continue;
            }

            if (children.Current is not IElement element || Skipped.Contains(element.LocalName))
                continue;

            if (Inline.Contains(element.LocalName))
            {
                loose.Append(element.LocalName.Equals("br", StringComparison.OrdinalIgnoreCase) ? "\n" : element.TextContent);
                continue;
            }

            Flush(blocks, loose);

            if (!IsReadAsABlock(element, blocks))
                open.Push(element.ChildNodes.GetEnumerator());
        }
    }

    /// <summary>Adds the block an element is, or answers false for a container whose children are walked instead.</summary>
    private static bool IsReadAsABlock(IElement element, List<Block> blocks)
    {
        switch (element.LocalName.ToLowerInvariant())
        {
            case "h1" or "h2" or "h3" or "h4" or "h5" or "h6":
                blocks.Add(new Heading(element.LocalName[1] - '0', Collapsed(element.TextContent)));
                return true;
            case "p" or "blockquote" or "dt" or "dd" or "figcaption" or "caption":
                blocks.Add(new Paragraph(Collapsed(element.TextContent)));
                return true;
            case "pre":
                blocks.Add(new CodeBlock(string.Empty, element.TextContent));
                return true;
            case "table":
                blocks.Add(new TableBlock(element.QuerySelector("caption")?.TextContent.Trim(), Rows(element)));
                return true;
            case "ul" or "ol":
                Items(element, blocks);
                return true;
            default:
                return false;
        }
    }

    /// <summary>The items of a list and of the lists nested in them, in document order, without recursion.</summary>
    private static void Items(IElement root, List<Block> blocks)
    {
        Stack<(IEnumerator<IElement> Items, int Depth, bool IsOrdered)> open = [];
        open.Push((Listed(root), 0, IsOrdered(root)));

        while (open.TryPeek(out var list))
        {
            if (!list.Items.MoveNext())
            {
                open.Pop();
                continue;
            }

            var item = list.Items.Current;
            blocks.Add(new ListItem(list.Depth, OwnText(item), list.IsOrdered));

            var nested = item.Children.Where(IsList).Reverse();

            foreach (var inner in nested)
                open.Push((Listed(inner), list.Depth + 1, IsOrdered(inner)));
        }
    }

    /// <summary>The text of a list item without the lists nested in it, which are items of their own.</summary>
    private static string OwnText(IElement item)
    {
        var own = item.ChildNodes.Where(node => node is not IElement element || !IsList(element));

        return Collapsed(string.Concat(own.Select(node => node.TextContent)));
    }

    private static IEnumerator<IElement> Listed(IElement list) =>
        list.Children.Where(child => child.LocalName.Equals("li", StringComparison.OrdinalIgnoreCase)).GetEnumerator();

    private static bool IsList(IElement element) => element.LocalName is "ul" or "ol";

    private static bool IsOrdered(IElement list) => list.LocalName.Equals("ol", StringComparison.OrdinalIgnoreCase);

    private static IReadOnlyList<IReadOnlyList<string>> Rows(IElement table) =>
    [
        .. table.QuerySelectorAll("tr")
            .Where(row => row.Closest("table") == table)
            .Select(row => (IReadOnlyList<string>)[.. row.Children.Where(cell => cell.LocalName is "td" or "th").Select(cell => Collapsed(cell.TextContent))])
            .Where(row => row.Count > 0)
    ];

    private static void Flush(List<Block> blocks, StringBuilder loose)
    {
        if (Collapsed(loose.ToString()) is { Length: > 0 } text)
            blocks.Add(new Paragraph(text));

        loose.Clear();
    }

    private static string Collapsed(string text) =>
        string.Join('\n', text.Split('\n').Select(line => string.Join(' ', line.Split((char[])[' ', '\t', '\r', ' '], StringSplitOptions.RemoveEmptyEntries))).Where(line => line.Length > 0));
}
