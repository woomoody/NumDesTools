using System.Collections.Concurrent;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MiniExcelLibs;
using NumDesTools.Config;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;
using ExcelReference = ExcelDna.Integration.ExcelReference;

// ReSharper disable All

#pragma warning disable CA1416

namespace NumDesTools;

public static partial class PubMetToExcel
{
    //查找资源文件
    public static Dictionary<string, List<string>> FindResourceFile(
        Dictionary<string, List<string>> longNumbers,
        string searchFolder
    )
    {
        // 线程安全字典：存储并行查找的结果
        var tempDict = new ConcurrentDictionary<string, List<string>>();

        var searchOptions = new EnumerationOptions
        {
            MatchCasing = MatchCasing.CaseInsensitive,
            RecurseSubdirectories = true,
        };

        // **并行遍历字典的 Key-Value 对**
        Parallel.ForEach(
            longNumbers,
            kvp =>
            {
                string dictKey = kvp.Key; // 原始 Key
                List<string> values = kvp.Value; // 该 Key 关联的 List<string>

                if (values.Count < 3)
                    return; // 确保 values 至少有 3 个元素

                string value1 = values[0]; // 第 1 个值
                string subFolder = values[1]; // 用于拼接路径
                string searchNum = values[2]; // 作为图片名称查找

                string newSearchPath = Path.Combine(searchFolder, subFolder); // 拼接路径
                List<string> foundImages = new List<string>();

                if (Directory.Exists(newSearchPath)) // 确保目录存在
                {
                    var files = Directory.EnumerateFiles(
                        newSearchPath,
                        $"{searchNum}.png",
                        searchOptions
                    );
                    foundImages.AddRange(files);
                }

                // **存入 Key，包含 (value1, searchNum, 图片路径)**
                if (foundImages.Count > 0)
                {
                    tempDict[dictKey] = new List<string> { value1, searchNum };
                    tempDict[dictKey].AddRange(foundImages); // 添加所有找到的图片路径
                }
                else
                {
                    tempDict[dictKey] = new List<string> { value1, searchNum }; // 即使没找到，也存储基础数据
                }
            }
        );

        // **保证返回的 Dictionary 顺序与 longNumbers 一致**
        var orderedDict = new Dictionary<string, List<string>>();
        foreach (var key in longNumbers.Keys)
        {
            if (tempDict.TryGetValue(key, out var value))
            {
                orderedDict[key] = value;
            }
            else
            {
                orderedDict[key] = new List<string> { longNumbers[key][0], longNumbers[key][2] }; // 确保所有 Key 都存在
            }
        }

        return orderedDict;
    }

    // 检查Excel单元格值是否重复
    public static List<(string, int, int, string, string, string)> CheckRepeatValue(
        string wkFullPath
    )
    {
        var sourceData = new List<(string, int, int, string, string, string)>();

        if (wkFullPath.Contains("#"))
        {
            return sourceData;
        }

        var sheetNames = MiniExcel.GetSheetNames(wkFullPath);

        foreach (var sheetName in sheetNames)
        {
            if (sheetName.Contains("#") || sheetName.Contains("Chart"))
                continue;
            var rows = MiniExcel
                .Query(
                    wkFullPath,
                    sheetName: sheetName,
                    configuration: NumDesAddIn.OnOffMiniExcelCatches
                )
                .ToList();

            if (rows.Count <= 4)
            {
                continue;
            }

            var dataRows = rows.Skip(3).ToList();

            if (dataRows.Count == 0)
            {
                continue;
            }

            // 检查第 1、2 列第 1 行的值是否为特定字符串，如果是则跳过该工作表
            if (
                dataRows.Any() && ((IDictionary<string, object>)dataRows[0])["A"]?.ToString() != "#"
                || ((IDictionary<string, object>)dataRows[0])["B"]?.ToString() == null
            )
            {
                continue;
            }

            // 检查 List 中第 2 列 是否有重复值，并返回重复值的行列号
            var duplicates = dataRows
                .Select((row, index) => new { Row = row, Index = index + 4 }) // 保留行号，+5 是因为跳过了前 4 行
                .Where(x => ((IDictionary<string, object>)x.Row)["B"] != null) // 忽略 null 值
                .GroupBy(x => ((IDictionary<string, object>)x.Row)["B"]) // 按第 2 列的值分组
                .Where(group => group.Count() > 1) // 找出重复值
                .SelectMany(group => group) // 展开分组
                .ToList();

            //转换数据格式
            foreach (var duplicate in duplicates)
            {
                var cellValue = ((IDictionary<string, object>)duplicate.Row)["B"].ToString();
                var cellRow = duplicate.Index;
                var cellCol = 2; // 第 2 列
                sourceData.Add((cellValue, cellRow, cellCol, sheetName, "数据重复", wkFullPath));
            }
        }

        return sourceData;
    }

    // 检查Excel单元格值的合法性
    public static List<(string, int, int, string, string, string)> ExcelCellValueFormatCheck(
        string cellValue,
        string typeCell,
        string sheetName,
        string filePath,
        int rowIndex,
        int colIndex
    )
    {
        var config = new GlobalVariable();
        var normalCharactersCheck = config.NormaKeyList;
        var specialCharactersCheck = config.SpecialKeyList;
        var coupleCharactersCheck = config.CoupleKeyList;

        var sourceData = new List<(string, int, int, string, string, string)>();

        if (cellValue != null)
        {
            if (
                normalCharactersCheck.Any(c => cellValue.Contains(c))
                && !typeCell.Contains("string")
            )
            {
                sourceData.Add(
                    (cellValue, rowIndex + 1, colIndex + 1, sheetName, "多逗号或中文逗号", filePath)
                );
            }

            if (
                specialCharactersCheck.Any(c => cellValue.Contains(c))
                && !typeCell.Contains("string")
            )
            {
                sourceData.Add(
                    (cellValue, rowIndex + 1, colIndex + 1, sheetName, "少逗号", filePath)
                );
            }

            foreach (var (leftString, rightString) in coupleCharactersCheck)
            {
                var leftStringCount = Regex
                    .Matches(cellValue, Regex.Escape(leftString), RegexOptions.IgnoreCase)
                    .Count;
                var RightStringCount = Regex
                    .Matches(cellValue, Regex.Escape(rightString), RegexOptions.IgnoreCase)
                    .Count;
                if (leftStringCount != RightStringCount)
                {
                    sourceData.Add(
                        (cellValue, rowIndex + 1, colIndex + 1, sheetName, "括号问题", filePath)
                    );
                    break;
                }

                if (leftString == "\"")
                {
                    int isDouble = leftStringCount % 2;
                    if (isDouble != 0)
                    {
                        sourceData.Add(
                            (
                                cellValue,
                                rowIndex + 1,
                                colIndex + 1,
                                sheetName,
                                "双引号问题",
                                filePath
                            )
                        );
                        break;
                    }
                }
            }
        }

        return sourceData;
    }

    // 判断某个工作表是否需要做公式检测。
    // 规则：只对 ...\Public\Excels\Tables\ 目录下的合法配置表做检测；
    //   - 工作簿文件名含 # → 非配置文件，跳过检测
    //   - Sheet 名含 #     → 辅助/非配置 Sheet，跳过检测
    //   - 工作簿文件名含 $ → 多 Sheet 配置工作簿，所有 Sheet 均为合法配置，需检测
    //   - 其余             → 仅 Sheet1 是合法配置 Sheet，其他 Sheet 跳过检测
    public static bool ShouldCheckFormula(string workbookFilePath, string sheetName)
    {
        if (
            !workbookFilePath.Contains(
                @"\Public\Excels\Tables\",
                StringComparison.OrdinalIgnoreCase
            )
        )
            return false;

        var fileName = Path.GetFileNameWithoutExtension(workbookFilePath);
        if (fileName.Contains('#'))
            return false;
        if (sheetName.Contains('#'))
            return false;

        if (fileName.Contains('$'))
            return true;

        return string.Equals(sheetName, "Sheet1", StringComparison.OrdinalIgnoreCase);
    }
}
