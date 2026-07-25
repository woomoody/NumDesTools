using System.Collections;
using System.Reflection;
using OfficeOpenXml;

namespace NumDesTools.Tests;

public sealed class OoxmlLazyReaderTests
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";
    private static readonly Type ReaderType = LoadReaderType();

    static OoxmlLazyReaderTests() => ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    [Fact]
    public void ReadDimension_ReturnsCorrectRange()
    {
        var sheetName = ReadSheetNames()[0];

        var dimension = Invoke("ReadDimension", ItemPath, sheetName);

        Assert.True(ReadProperty<int>(dimension, "Item1") > 60_000);
        Assert.True(ReadProperty<int>(dimension, "Item2") > 0);
    }

    [Fact]
    public void ReadRows_First200_ReturnsExpectedContent()
    {
        var rows = ReadRows(ReadSheetNames()[0], 200, 0);

        Assert.Equal(200, rows.Count);
        Assert.All(
            rows,
            row =>
            {
                Assert.True(ReadProperty<int>(row, "RowNum") > 0);
                Assert.NotEmpty(ReadCells(row));
            }
        );
    }

    [Fact]
    public void ReadRows_SharedStringDecoded()
    {
        var sheetName = ReadSheetNames()[0];
        var rows = ReadRows(sheetName, 200, 0);
        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var sheet = package.Workbook.Worksheets[sheetName];
        var samples = rows.SelectMany(row =>
                ReadCells(row)
                    .Where(cell => !string.IsNullOrEmpty(cell.Value))
                    .Select(cell => (Row: ReadProperty<int>(row, "RowNum"), cell.Key, cell.Value))
            )
            .Where(sample =>
                string.Equals(
                    sheet.Cells[sample.Row, GetColumnNumber(sample.Key)].Value?.ToString(),
                    sample.Value,
                    StringComparison.Ordinal
                )
            )
            .Take(5)
            .ToList();

        Assert.Equal(5, samples.Count);
        Assert.All(
            samples,
            sample =>
                Assert.Equal(
                    sheet.Cells[sample.Row, GetColumnNumber(sample.Key)].Value?.ToString(),
                    sample.Value
                )
        );
    }

    [Fact]
    public void ReadRows_SkipRows_ContinuesCorrectly()
    {
        var rows = ReadRows(ReadSheetNames()[0], 100, 200);

        Assert.Equal(100, rows.Count);
        Assert.Equal(201, ReadProperty<int>(rows[0], "RowNum"));
    }

    [Fact]
    public void ReadRows_NonexistentSheet_ReturnsEmpty()
    {
        var rows = ReadRows("Missing sheet", 200, 0);

        Assert.Empty(rows);
    }

    [Fact]
    public void ReadSheetNames_ReturnsAllSheets()
    {
        var names = ReadSheetNames();

        Assert.NotEmpty(names);
        Assert.False(string.IsNullOrWhiteSpace(names[0]));
    }

    private static List<string> ReadSheetNames() =>
        Assert.IsType<List<string>>(Invoke("ReadSheetNames", ItemPath));

    private static List<object> ReadRows(string sheetName, int maxRows, int skipRows) =>
        Assert
            .IsAssignableFrom<IEnumerable>(
                Invoke("ReadRows", ItemPath, sheetName, maxRows, skipRows)
            )
            .Cast<object>()
            .ToList();

    private static Dictionary<string, string> ReadCells(object row) =>
        Assert.IsType<Dictionary<string, string>>(ReadProperty<object>(row, "Cells"));

    private static object Invoke(string methodName, params object[] arguments) =>
        ReaderType
            .GetMethod(methodName, BindingFlags.Public | BindingFlags.Static)!
            .Invoke(null, arguments)!;

    private static T ReadProperty<T>(object target, string propertyName)
    {
        var memberValue =
            target.GetType().GetProperty(propertyName)?.GetValue(target)
            ?? target.GetType().GetField(propertyName)?.GetValue(target);
        return Assert.IsAssignableFrom<T>(memberValue);
    }

    private static Type LoadReaderType()
    {
        var assemblyPath = Path.GetFullPath(
            Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "NumDesTools.XlsxEditor",
                "bin",
                "Debug",
                "net9.0-windows",
                "NumDesTools.XlsxEditor.dll"
            )
        );
        var assembly = Assembly.LoadFrom(assemblyPath);
        return assembly.GetType("NumDesTools.XlsxEditor.OoxmlLazyReader", throwOnError: true)!;
    }

    private static int GetColumnNumber(string columnName)
    {
        var result = 0;
        foreach (var character in columnName)
        {
            result = result * 26 + character - 'A' + 1;
        }

        return result;
    }
}
