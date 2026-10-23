using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace Snail.Toolkit.Documents.Tests.Chaos;

/// <summary>Files and streams built to find where a promise of the library stops holding.</summary>
internal static class Hostility
{
    /// <summary>A workbook of <paramref name="rows"/> rows of one cell each, stored uncompressed so the zip limits let it through.</summary>
    /// <remarks>
    /// Written by hand rather than through the SDK: a million <c>Row</c> objects would cost the test more than the reader, and the
    /// point is the reader's cost. Stored, the file is as large as its XML, so the unpacking ratio is 1 and the bytes are under the limit.
    /// </remarks>
    public static byte[] TallWorkbook(int rows)
    {
        using var buffer = new MemoryStream();

        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");
            Write(archive, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
            Write(archive, "xl/workbook.xml", """<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""");
            Write(archive, "xl/_rels/workbook.xml.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");

            using var sheet = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.NoCompression).Open(), new UTF8Encoding(false));
            sheet.Write("""<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>""");

            for (int row = 1; row <= rows; row++)
                sheet.Write($"<row r=\"{row}\"><c r=\"A{row}\"><v>1</v></c></row>");

            sheet.Write("</sheetData></worksheet>");
        }

        return buffer.ToArray();
    }

    /// <summary>A Word document of <paramref name="paragraphs"/> short paragraphs, stored uncompressed so the zip limits let it through.</summary>
    public static byte[] TallDocument(int paragraphs) =>
        WordDocument(body =>
        {
            for (int paragraph = 1; paragraph <= paragraphs; paragraph++)
                body.Write($"<w:p><w:r><w:t>Paragraph {paragraph} of the report says little.</w:t></w:r></w:p>");
        });

    /// <summary>A Word document of one paragraph of <paramref name="runs"/> runs of fifty thousand letters each, stored uncompressed.</summary>
    public static byte[] WideParagraph(int runs) =>
        WordDocument(body =>
        {
            string letters = new('x', 50_000);
            body.Write("<w:p>");

            for (int run = 0; run < runs; run++)
                body.Write($"<w:r><w:t>{letters}</w:t></w:r>");

            body.Write("</w:p>");
        });

    /// <summary>A Word document of one table of <paramref name="rows"/> rows of one short cell, stored uncompressed.</summary>
    public static byte[] TallTable(int rows) =>
        WordDocument(body =>
        {
            body.Write("<w:tbl>");

            for (int row = 1; row <= rows; row++)
                body.Write($"<w:tr><w:tc><w:p><w:r><w:t>row {row}</w:t></w:r></w:p></w:tc></w:tr>");

            body.Write("</w:tbl>");
        });

    /// <summary>A presentation of one slide whose one shape holds <paramref name="paragraphs"/> paragraphs of fifty thousand letters each, stored uncompressed.</summary>
    public static byte[] WideSlide(int paragraphs)
    {
        using var buffer = new MemoryStream();

        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>""");
            Write(archive, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>""");
            Write(archive, "ppt/presentation.xml", """<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>""");
            Write(archive, "ppt/_rels/presentation.xml.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""");

            using var slide = new StreamWriter(archive.CreateEntry("ppt/slides/slide1.xml", CompressionLevel.NoCompression).Open(), new UTF8Encoding(false));
            slide.Write("""<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>""");
            string letters = new('x', 50_000);

            for (int paragraph = 0; paragraph < paragraphs; paragraph++)
                slide.Write($"<a:p><a:r><a:t>{letters}</a:t></a:r></a:p>");

            slide.Write("</p:txBody></p:sp></p:spTree></p:cSld></p:sld>");
        }

        return buffer.ToArray();
    }

    private static byte[] WordDocument(Action<StreamWriter> content)
    {
        using var buffer = new MemoryStream();

        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""");
            Write(archive, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""");

            using var body = new StreamWriter(archive.CreateEntry("word/document.xml", CompressionLevel.NoCompression).Open(), new UTF8Encoding(false));
            body.Write("""<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>""");
            content(body);
            body.Write("</w:body></w:document>");
        }

        return buffer.ToArray();
    }

    /// <summary>A workbook of one empty sheet whose styles hold <paramref name="formats"/> cell formats, each showing a date; stored uncompressed.</summary>
    public static byte[] BloatedStyles(int formats) =>
        Workbook(
            """<sheet name="Sheet1" sheetId="1" r:id="rId2"/>""",
            styles =>
            {
                styles.Write("""<numFmts count="1"><numFmt numFmtId="164" formatCode="dd.mm.yyyy"/></numFmts><cellXfs>""");

                for (int format = 0; format < formats; format++)
                    styles.Write("""<xf numFmtId="164"/>""");

                styles.Write("</cellXfs>");
            });

    /// <summary>A workbook whose list of sheets names <paramref name="sheets"/> sheets, every one under a relationship to the styles part; stored uncompressed.</summary>
    /// <remarks>Pointing at a part that is not a sheet keeps the reader from reading anything but the list: what is measured is the list.</remarks>
    public static byte[] ListOfSheets(int sheets)
    {
        var list = new StringBuilder();

        for (int sheet = 1; sheet <= sheets; sheet++)
            list.Append($"<sheet name=\"S{sheet}\" sheetId=\"{sheet}\" r:id=\"rId1\"/>");

        return Workbook(list.ToString(), _ => { });
    }

    /// <summary>A presentation whose list of slides names <paramref name="slides"/> slides, every one under a relationship to its properties part; stored uncompressed.</summary>
    public static byte[] ListOfSlides(int slides)
    {
        using var buffer = new MemoryStream();

        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/></Types>""");
            Write(archive, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>""");
            Write(archive, "ppt/_rels/presentation.xml.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/></Relationships>""");
            Write(archive, "ppt/presProps.xml", """<?xml version="1.0" encoding="UTF-8"?><p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>""");

            using var presentation = new StreamWriter(archive.CreateEntry("ppt/presentation.xml", CompressionLevel.NoCompression).Open(), new UTF8Encoding(false));
            presentation.Write("""<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>""");

            for (int slide = 0; slide < slides; slide++)
                presentation.Write($"<p:sldId id=\"{256 + slide}\" r:id=\"rId1\"/>");

            presentation.Write("</p:sldIdLst></p:presentation>");
        }

        return buffer.ToArray();
    }

    /// <summary>A workbook of the sheets <paramref name="sheets"/> lists, with a styles part <paramref name="styles"/> fills and one empty worksheet; stored uncompressed.</summary>
    private static byte[] Workbook(string sheets, Action<StreamWriter> styles)
    {
        using var buffer = new MemoryStream();

        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");
            Write(archive, "_rels/.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
            Write(archive, "xl/_rels/workbook.xml.rels", """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");
            Write(archive, "xl/worksheets/sheet1.xml", """<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="0"><v>46305</v></c></row></sheetData></worksheet>""");

            using (var workbook = new StreamWriter(archive.CreateEntry("xl/workbook.xml", CompressionLevel.NoCompression).Open(), new UTF8Encoding(false)))
            {
                workbook.Write("""<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>""");
                workbook.Write(sheets);
                workbook.Write("</sheets></workbook>");
            }

            using var stylesheet = new StreamWriter(archive.CreateEntry("xl/styles.xml", CompressionLevel.NoCompression).Open(), new UTF8Encoding(false));
            stylesheet.Write("""<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
            styles(stylesheet);
            stylesheet.Write("</styleSheet>");
        }

        return buffer.ToArray();
    }

    /// <summary>A zip holding sixty-four megabytes of zeros whose directory says the entry unpacks to a thousand bytes.</summary>
    /// <remarks>The sizes in a zip's directory are claims, and a bomb claims whatever lets it through; the deflate stream inside holds the truth.</remarks>
    public static byte[] LyingZip()
    {
        using var buffer = new MemoryStream();

        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        using (var entry = archive.CreateEntry("word/document.xml", CompressionLevel.Optimal).Open())
            entry.Write(new byte[64 * 1024 * 1024]);

        byte[] zip = buffer.ToArray();
        Patch(zip, [0x50, 0x4B, 0x03, 0x04], 22);
        Patch(zip, [0x50, 0x4B, 0x01, 0x02], 24);

        return zip;
    }

    /// <summary>Text of about <paramref name="megabytes"/> megabytes in short paragraphs, with Windows line endings so nothing is read in place.</summary>
    public static byte[] Prose(int megabytes)
    {
        var text = new StringBuilder(megabytes * 1024 * 1024 + 128);
        string line = $"{new string('x', 78)}\r\n\r\n";

        while (text.Length < megabytes * 1024 * 1024)
            text.Append(line);

        return Encoding.ASCII.GetBytes(text.ToString());
    }

    /// <summary>Two thousand rows that each hold one cell at XFD, the last column: eleven kilobytes that pad to thirty-two million cells.</summary>
    public static byte[] SparseWorkbook()
    {
        using var stream = new MemoryStream();

        using (var book = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbook = book.AddWorkbookPart();
            workbook.Workbook = new S.Workbook();
            var data = new S.SheetData();

            for (uint row = 1; row <= 2_000; row++)
                data.Append(new S.Row(new S.Cell { CellReference = row == 1 ? $"A{row}" : $"XFD{row}", CellValue = new S.CellValue("1") }) { RowIndex = row });

            var part = workbook.AddNewPart<WorksheetPart>();
            part.Worksheet = new S.Worksheet(data);
            workbook.Workbook.AppendChild(new S.Sheets(new S.Sheet { Id = workbook.GetIdOfPart(part), SheetId = 1, Name = "Sheet 1" }));
        }

        return stream.ToArray();
    }

    private static void Write(ZipArchive archive, string name, string xml)
    {
        using var entry = archive.CreateEntry(name).Open();
        entry.Write(Encoding.UTF8.GetBytes(xml));
    }

    private static void Patch(byte[] zip, byte[] signature, int offset)
    {
        int at = zip.AsSpan().IndexOf(signature);
        BinaryPrimitives.WriteUInt32LittleEndian(zip.AsSpan(at + offset), 1000);
    }
}

/// <summary>The tests that measure the heap of the whole process, run while nothing else does.</summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class PeakHeapCollection
{
    public const string Name = "Peak heap";
}

/// <summary>The most the managed heap held while something ran, sampled every few milliseconds: what a container must have, which the bytes allocated in total do not say.</summary>
internal sealed class PeakHeap : IDisposable
{
    private readonly Thread _sampling;

    private volatile bool _isStopped;

    private long _peak;

    public PeakHeap()
    {
        _sampling = new Thread(Sample) { IsBackground = true };
        _sampling.Start();
    }

    /// <summary>The largest heap seen, in bytes.</summary>
    public long Bytes
    {
        get
        {
            _isStopped = true;
            _sampling.Join();

            return Math.Max(_peak, GC.GetTotalMemory(forceFullCollection: false));
        }
    }

    public void Dispose() => _isStopped = true;

    private void Sample()
    {
        while (!_isStopped)
        {
            _peak = Math.Max(_peak, GC.GetTotalMemory(forceFullCollection: false));
            Thread.Sleep(5);
        }
    }
}

/// <summary>A stream read once from the start, as a network or a pipe gives it.</summary>
internal sealed class ForwardOnlyStream(byte[] bytes) : MemoryStream(bytes)
{
    public override bool CanSeek => false;
}
