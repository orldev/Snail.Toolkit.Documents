using DocumentFormat.OpenXml;

namespace Snail.Toolkit.Documents.Readers;

/// <summary>What the events of an Office part reader say, shared by the readers that walk a part element by element.</summary>
internal static class PartEvents
{
    extension(OpenXmlReader reader)
    {
        /// <summary>Whether the reader stands on the end of the element opened at <paramref name="depth"/>.</summary>
        /// <remarks>An empty element has an end event too, so a walk that stops here never runs past the element it began in.</remarks>
        public bool IsEndAt(int depth) =>
            reader.IsEndElement && reader.Depth == depth;
    }
}
