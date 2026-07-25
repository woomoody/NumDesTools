using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace NumDesTools.XlsxEditor;

public sealed record RawRow(
    int RowNum,
    Dictionary<string, string> Cells,
    Dictionary<(int Row, int Col), string> Comments
);

internal static class OoxmlLazyReader
{
    private const string WorkbookPath = "xl/workbook.xml";
    private const string WorkbookRelationshipsPath = "xl/_rels/workbook.xml.rels";
    private const string SharedStringsPath = "xl/sharedStrings.xml";
    private const string RelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static (int Rows, int Cols) ReadDimension(string xlsxPath, string sheetName)
    {
        if (ShouldSkip(sheetName))
        {
            return (0, 0);
        }

        using var archive = OpenArchive(xlsxPath);
        var sheetPath = FindSheetPath(archive, sheetName);
        var entry = sheetPath is null ? null : archive.GetEntry(sheetPath);
        if (entry is null)
        {
            return (0, 0);
        }

        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            if (reader is not { NodeType: XmlNodeType.Element, LocalName: "dimension" })
            {
                continue;
            }

            return ParseDimension(reader.GetAttribute("ref"));
        }

        return (0, 0);
    }

    public static IEnumerable<RawRow> ReadRows(
        string xlsxPath,
        string sheetName,
        int maxRows = 200,
        int skipRows = 0
    )
    {
        ArgumentOutOfRangeException.ThrowIfNegative(maxRows);
        ArgumentOutOfRangeException.ThrowIfNegative(skipRows);

        if (maxRows is 0 || ShouldSkip(sheetName))
        {
            yield break;
        }

        using var archive = OpenArchive(xlsxPath);
        var sheetPath = FindSheetPath(archive, sheetName);
        var entry = sheetPath is null ? null : archive.GetEntry(sheetPath);
        if (entry is null)
        {
            yield break;
        }

        var sharedStrings = ReadSharedStrings(archive);
        var comments = ReadComments(archive, sheetPath!);
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        var rowsSeen = 0;
        var rowsReturned = 0;

        while (rowsReturned < maxRows && reader.Read())
        {
            if (reader is not { NodeType: XmlNodeType.Element, LocalName: "row" })
            {
                continue;
            }

            rowsSeen++;
            var row = ReadRow(reader, sharedStrings, comments);
            if (rowsSeen <= skipRows)
            {
                continue;
            }

            rowsReturned++;
            yield return row;
        }
    }

    public static List<string> ReadSheetNames(string xlsxPath)
    {
        using var archive = OpenArchive(xlsxPath);
        var entry = archive.GetEntry(WorkbookPath);
        if (entry is null)
        {
            return [];
        }

        var names = new List<string>();
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            if (reader is not { NodeType: XmlNodeType.Element, LocalName: "sheet" })
            {
                continue;
            }

            var name = reader.GetAttribute("name");
            if (!ShouldSkip(name))
            {
                names.Add(name!);
            }
        }

        return names;
    }

    private static RawRow ReadRow(
        XmlReader reader,
        IReadOnlyList<string> sharedStrings,
        IReadOnlyDictionary<(int Row, int Col), string> comments
    )
    {
        var rowNumber = ParsePositiveInteger(reader.GetAttribute("r"));
        var cells = new Dictionary<string, string>(StringComparer.Ordinal);
        var rowComments = new Dictionary<(int Row, int Col), string>();

        if (!reader.IsEmptyElement)
        {
            using var rowReader = reader.ReadSubtree();
            while (rowReader.Read())
            {
                if (rowReader is not { NodeType: XmlNodeType.Element, LocalName: "c" })
                {
                    continue;
                }

                var reference = rowReader.GetAttribute("r");
                var columnName = GetColumnName(reference);
                if (columnName.Length is 0)
                {
                    continue;
                }

                var type = rowReader.GetAttribute("t");
                var value = ReadCellValue(rowReader, type, sharedStrings);
                cells[columnName] = value;

                var columnNumber = GetColumnNumber(columnName);
                if (comments.TryGetValue((rowNumber, columnNumber), out var comment))
                {
                    rowComments[(rowNumber, columnNumber)] = comment;
                }
            }
        }

        return new RawRow(rowNumber, cells, rowComments);
    }

    private static string ReadCellValue(
        XmlReader reader,
        string? type,
        IReadOnlyList<string> sharedStrings
    )
    {
        if (reader.IsEmptyElement)
        {
            return string.Empty;
        }

        using var cellReader = reader.ReadSubtree();
        while (cellReader.Read())
        {
            if (
                cellReader.NodeType is not XmlNodeType.Element
                || cellReader.LocalName is not ("v" or "t")
            )
            {
                continue;
            }

            var value = cellReader.ReadElementContentAsString();
            if (
                type is "s"
                && int.TryParse(
                    value,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var index
                )
                && index >= 0
                && index < sharedStrings.Count
            )
            {
                return sharedStrings[index];
            }

            return value;
        }

        return string.Empty;
    }

    private static List<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry(SharedStringsPath);
        if (entry is null)
        {
            return [];
        }

        var result = new List<string>();
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            if (reader is not { NodeType: XmlNodeType.Element, LocalName: "si" })
            {
                continue;
            }

            result.Add(ReadConcatenatedText(reader));
        }

        return result;
    }

    private static Dictionary<(int Row, int Col), string> ReadComments(
        ZipArchive archive,
        string sheetPath
    )
    {
        var relationshipsPath = GetRelationshipsPath(sheetPath);
        var relationshipsEntry = archive.GetEntry(relationshipsPath);
        if (relationshipsEntry is null)
        {
            return [];
        }

        string? commentsPath = null;
        using (var stream = relationshipsEntry.Open())
        using (var reader = CreateReader(stream))
        {
            while (reader.Read())
            {
                var relationshipType = reader.GetAttribute("Type");
                if (
                    reader is not { NodeType: XmlNodeType.Element, LocalName: "Relationship" }
                    || relationshipType is null
                    || !relationshipType.EndsWith("/comments", StringComparison.Ordinal)
                )
                {
                    continue;
                }

                commentsPath = ResolvePartPath(sheetPath, reader.GetAttribute("Target"));
                break;
            }
        }

        var commentsEntry = commentsPath is null ? null : archive.GetEntry(commentsPath);
        if (commentsEntry is null)
        {
            return [];
        }

        var result = new Dictionary<(int Row, int Col), string>();
        using var commentsStream = commentsEntry.Open();
        using var commentsReader = CreateReader(commentsStream);
        while (commentsReader.Read())
        {
            if (commentsReader is not { NodeType: XmlNodeType.Element, LocalName: "comment" })
            {
                continue;
            }

            var reference = commentsReader.GetAttribute("ref");
            var row = GetRowNumber(reference);
            var column = GetColumnNumber(GetColumnName(reference));
            if (row > 0 && column > 0)
            {
                result[(row, column)] = ReadConcatenatedText(commentsReader);
            }
        }

        return result;
    }

    private static string ReadConcatenatedText(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            return string.Empty;
        }

        var text = new System.Text.StringBuilder();
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            if (subtree is { NodeType: XmlNodeType.Element, LocalName: "t" })
            {
                text.Append(subtree.ReadElementContentAsString());
            }
        }

        return text.ToString();
    }

    private static string? FindSheetPath(ZipArchive archive, string sheetName)
    {
        var relationshipId = FindSheetRelationshipId(archive, sheetName);
        if (relationshipId is null)
        {
            return null;
        }

        var relationshipsEntry = archive.GetEntry(WorkbookRelationshipsPath);
        if (relationshipsEntry is null)
        {
            return null;
        }

        using var stream = relationshipsEntry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            if (
                reader is not { NodeType: XmlNodeType.Element, LocalName: "Relationship" }
                || reader.GetAttribute("Id") != relationshipId
            )
            {
                continue;
            }

            return ResolvePartPath(WorkbookPath, reader.GetAttribute("Target"));
        }

        return null;
    }

    private static string? FindSheetRelationshipId(ZipArchive archive, string sheetName)
    {
        var workbookEntry = archive.GetEntry(WorkbookPath);
        if (workbookEntry is null)
        {
            return null;
        }

        using var stream = workbookEntry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            if (
                reader is not { NodeType: XmlNodeType.Element, LocalName: "sheet" }
                || !string.Equals(reader.GetAttribute("name"), sheetName, StringComparison.Ordinal)
            )
            {
                continue;
            }

            return reader.GetAttribute("id", RelationshipNamespace);
        }

        return null;
    }

    private static (int Rows, int Cols) ParseDimension(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return (0, 0);
        }

        var endReference = reference.Split(':', 2)[^1];
        return (GetRowNumber(endReference), GetColumnNumber(GetColumnName(endReference)));
    }

    private static string GetColumnName(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
        {
            return string.Empty;
        }

        var length = 0;
        while (length < reference.Length && char.IsLetter(reference[length]))
        {
            length++;
        }

        return reference[..length].ToUpperInvariant();
    }

    private static int GetRowNumber(string? reference) =>
        reference is null
            ? 0
            : ParsePositiveInteger(reference.AsSpan(GetColumnName(reference).Length));

    private static int GetColumnNumber(string columnName)
    {
        var result = 0;
        foreach (var character in columnName)
        {
            result = result * 26 + character - 'A' + 1;
        }

        return result;
    }

    private static int ParsePositiveInteger(string? value) =>
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var result)
            ? result
            : 0;

    private static int ParsePositiveInteger(ReadOnlySpan<char> value) =>
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var result)
            ? result
            : 0;

    private static string GetRelationshipsPath(string partPath)
    {
        var separator = partPath.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : partPath[..(separator + 1)];
        var fileName = separator < 0 ? partPath : partPath[(separator + 1)..];
        return $"{directory}_rels/{fileName}.rels";
    }

    private static string ResolvePartPath(string sourcePartPath, string? target)
    {
        if (string.IsNullOrWhiteSpace(target))
        {
            return string.Empty;
        }

        if (target.StartsWith('/'))
        {
            return target.TrimStart('/');
        }

        var separator = sourcePartPath.LastIndexOf('/');
        var directory = separator < 0 ? string.Empty : sourcePartPath[..(separator + 1)];
        var uri = new Uri(new Uri($"http://package/{directory}"), target);
        return Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/');
    }

    private static bool ShouldSkip(string? sheetName) =>
        string.IsNullOrEmpty(sheetName) || sheetName.StartsWith('#');

    private static XmlReader CreateReader(Stream stream) =>
        XmlReader.Create(
            stream,
            new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                CloseInput = false,
            }
        );

    private static ZipArchive OpenArchive(string xlsxPath)
    {
        try
        {
            return ZipFile.OpenRead(xlsxPath);
        }
        catch (InvalidDataException exception)
        {
            throw new IOException(
                $"The xlsx file is not a valid zip archive: {xlsxPath}",
                exception
            );
        }
    }
}
