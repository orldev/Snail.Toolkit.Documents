# Snail.Toolkit.Documents

Reads PDF, Word, Excel, PowerPoint, HTML, CSV, JSON, Markdown and text into one Markdown text, with its pages, headings
and tables — in .NET, with no Office, no LibreOffice and no network.

```csharp
var reading = await DocumentReader.Every.ReadAsync(File.OpenRead("invoice.pdf"), "invoice.pdf");

if (reading.IsRead)
    Console.WriteLine(reading.Document!.Text);
else
    Console.WriteLine($"{reading.Failure}: {reading.Reason}");
```

## What it reads

| Format | Extensions | Read as |
| --- | --- | --- |
| Word | `.docx` `.docm` | headings, paragraphs, lists, tables, content controls |
| Excel | `.xlsx` `.xlsm` `.xls` | a part per sheet, its used cells one table under the letters of the columns and the numbers of the rows |
| PowerPoint | `.pptx` `.pptm` | a part per slide: its title, text, tables and the speaker's notes |
| PDF | `.pdf` | a part per page, its text layer in reading order; a page without one is marked as a scan |
| HTML | `.html` `.htm` | headings, paragraphs, lists, tables; scripts and styles left out |
| CSV | `.csv` `.tsv` | one table, the separator taken from the first line |
| JSON | `.json` | a list of flat records as a table, anything else as indented JSON |
| Markdown, text | `.md` `.markdown` `.txt` `.text` `.log` | as written |

Text files are read as UTF-8, or by their byte order mark, and as Windows-1251 when they are not valid UTF-8 — the way
a Russian Excel saves a CSV.

## A document

`Document` is the whole text as Markdown and where everything in it stands: `Parts` (pages, sheets, slides, each with
where its text starts and ends), `Headings` and `Tables`. The positions are offsets into `Text`, taken in the same pass
that wrote it, so a page a caller finds in the outline is the page it reads.

A sheet's table keeps where its cells stood: `Table.Addresses.Of(row, column)` is `B7`, and the Markdown shows it.

```
## Sheet 1: Invoices

|  | B | C | D |
| --- | --- | --- | --- |
| 2 | Number | Date | Sum |
| 3 | INV-7 | 2026-10-14 | 118000.5 |
| 5 | INV-8 |  | 42 |
```

## Failures are answers

`ReadAsync` throws only when its caller cancels it. Every other way a file fails comes back as a `Reading` with the
reason in the reader's words:

| `ReadingFailure` | When |
| --- | --- |
| `UnknownFormat` | no reader takes a file of this name |
| `Damaged` | the file is damaged, or is not what its name says |
| `Locked` | the file is protected by a password |
| `TooLarge` | the file unpacks past `ReadingLimits`: 20 million characters, a million cells in one table, 256 MB unpacked or a hundred times its size for an Office file |
| `TooSlow` | reading took longer than 30 seconds |

The limits hold whatever the size of the file: a workbook of 11 kB can come to 98 million characters, and a zip of a
few kilobytes to gigabytes. Each limit is stated with the file that found it in `ReadingLimits`.

## Scanned pages

A PDF page without a text layer is a `Part` with `IsScanned`. Its picture is what a model that sees reads:

```csharp
int[] scans = [.. document.Parts.Where(part => part.IsScanned).Select(part => part.Number)];
IReadOnlyList<ScannedPage> pages = ScannedPages.Pictures(File.OpenRead("scan.pdf"), scans);
```

Each page comes back as a PNG or a JPEG, or without a picture when no managed decoder opens it (fax and JBIG2 scans).

## In a container

```csharp
services.AddDocumentReading();   // every IFormatReader, and the DocumentReader over them
```

A program that reads fewer formats registers its own `IFormatReader`s and a `DocumentReader` over them.

The public API: [docs/reference](docs/reference). Built on [Snail.Sdk](https://github.com/orldev/Snail.Sdk): the rules,
the public API kept in `src/PublicAPI.*.txt`, validation against the last release, the shared workflows.
