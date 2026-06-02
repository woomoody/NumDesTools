using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using static System.String;
using Match = System.Text.RegularExpressions.Match;

#pragma warning disable CA1416

namespace NumDesTools;

public partial class ExcelUdf
{
    private static readonly dynamic IndexWk = AppServices.App.ActiveWorkbook;
    private static readonly dynamic ExcelPath = IndexWk.Path;

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "寻找同层级指定表格字段所在列"
    )]
    public static int FindKeyCol(
        [ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标行")] int row,
        [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1"
    )
    {
        var path = ExcelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
        var rowSource = sheet.GetRow(row);
        for (int j = rowSource.FirstCellNum; j <= rowSource.LastCellNum; j++)
        {
            var cell = rowSource.GetCell(j);
            if (cell != null)
            {
                var cellValue = cell.ToString();
                if (cellValue == searchValue)
                {
                    workbook.Close();
                    fs.Close();
                    return j;
                }
            }
        }

        workbook.Close();
        fs.Close();
        return -1;
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "寻找同层级指定表格字段所在行"
    )]
    public static int FindKeyRow(
        [ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标列")] int col,
        [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1"
    )
    {
        var path = ExcelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
        for (var i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
        {
            var rowSource = sheet.GetRow(i);
            if (rowSource != null)
            {
                var cell = rowSource.GetCell(col);
                var cellValue = cell.ToString();
                if (cellValue == searchValue)
                {
                    workbook.Close();
                    fs.Close();
                    return i;
                }
            }
        }

        workbook.Close();
        fs.Close();
        return -1;
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "寻找同层级指定表格字段所在列指定行的值"
    )]
    public static string FindKeyColToRow(
        [ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标行")] int row,
        [ExcelArgument(Description = "输出目标行")] int rowOut,
        [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1"
    )
    {
        var path = ExcelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
        var rowSource = sheet.GetRow(row);
        for (int j = rowSource.FirstCellNum; j <= rowSource.LastCellNum; j++)
        {
            var cell = rowSource.GetCell(j);
            if (cell != null)
            {
                var cellValue = cell.ToString();
                if (cellValue == searchValue)
                {
                    var outRowSource = sheet.GetRow(rowOut);
                    var outCell = outRowSource.GetCell(j);
                    var outCellValue = outCell.ToString();
                    workbook.Close();
                    fs.Close();
                    return outCellValue;
                }
            }
        }

        workbook.Close();
        fs.Close();
        return "Error未找到";
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "寻找同层级指定表格字段所在行指定列的值"
    )]
    public static string FindKeyRowToCol(
        [ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标列")] int col,
        [ExcelArgument(Description = "输出目标列")] int outCol,
        [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1"
    )
    {
        var path = ExcelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
        for (var i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
        {
            var rowSource = sheet.GetRow(i);
            if (rowSource != null)
            {
                var cell = rowSource.GetCell(col);
                var cellValue = cell.ToString();
                if (cellValue == searchValue)
                {
                    var outCell = rowSource.GetCell(outCol);
                    var outCellValue = outCell?.ToString() ?? string.Empty;
                    workbook.Close();
                    fs.Close();
                    return outCellValue;
                }
            }
        }

        workbook.Close();
        fs.Close();
        return "Error未找到";
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "兼容索引，索引单元格有值则相对索引，否则绝对索引，索引最靠近的单元格（上-左）"
    )]
    public static string FindKeyClose(
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "单元格")]
            object inputRange,
        [ExcelArgument(Description = "行索引或列索引")] bool isRow
    )
    {
        if (inputRange is ExcelReference cellRef)
        {
            var sheet = AppServices.App.ActiveSheet;
            var rangeRow = cellRef.RowFirst + 1;
            var rangeCol = cellRef.ColumnFirst + 1;
            var rangeValue = sheet.Cells[rangeRow, rangeCol].Value;
            if (rangeValue == null)
            {
                if (isRow)
                {
                    var count = rangeRow;
                    while (count > 0)
                    {
                        var newRangeValue = sheet.Cells[count, rangeCol].Value;
                        if (newRangeValue == null)
                            count--;
                        else
                            return newRangeValue.ToString();
                    }
                }
                else
                {
                    var count = rangeCol;
                    while (count > 0)
                    {
                        var newRangeValue = sheet.Cells[rangeRow, count].Value;
                        if (newRangeValue == null)
                            count--;
                        else
                            return newRangeValue.ToString();
                    }
                }
            }
            else
            {
                return rangeValue.ToString();
            }
        }

        return "Error未找到";
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "替换正则到的数据为指定值"
    )]
    public static string ReplaceKey(
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "单元格")]
            string inputRange,
        [ExcelArgument(AllowReference = true, Description = "正则方案：%d", Name = "正则方案")]
            string regexMethod,
        [ExcelArgument(AllowReference = true, Description = "替换值：abc", Name = "替换值")]
            string replaceValue
    )
    {
        var result = Regex.Replace(inputRange, regexMethod, replaceValue);

        return result;
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "查找正则匹配到的数据并输出为字符串"
    )]
    public static string FindKey(
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "单元格")]
            string inputRange,
        [ExcelArgument(AllowReference = true, Description = "正则方案：%d", Name = "正则方案")]
            string regexMethod
    )
    {
        // 参数校验
        if (IsNullOrEmpty(inputRange))
            return "输入单元格地址不能为空";

        if (IsNullOrEmpty(regexMethod))
            regexMethod = @"\d+";

        try
        {
            // 使用正则表达式匹配
            var matches = Regex.Matches(inputRange, regexMethod);

            // 将匹配结果连接为字符串
            var result = Join(", ", matches.Select(m => m.Value));

            // 如果没有匹配到内容，返回提示信息
            if (IsNullOrEmpty(result))
                return "未找到匹配内容";

            return result;
        }
        catch (ArgumentException ex)
        {
            // 捕获正则表达式语法错误
            return $"正则表达式错误：{ex.Message}";
        }
        catch (Exception ex)
        {
            // 捕获其他异常
            return $"发生错误：{ex.Message}";
        }
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "替换指定位置匹配到的值为指定值"
    )]
    public static string ReplaceKeyByIndex(
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "单元格")]
            string inputRange,
        [ExcelArgument(AllowReference = true, Description = "匹配位置", Name = "匹配位置信息")]
            object[,] matchIndex,
        [ExcelArgument(AllowReference = true, Description = "替换值", Name = "替换值信息")]
            object[,] replaceValue,
        [ExcelArgument(AllowReference = true, Description = "正则方案", Name = "正则方案")]
            string regexMethod
    )
    {
        var rows = matchIndex.GetLength(0);
        var cols = matchIndex.GetLength(1);

        var replacements = new Dictionary<int, string>();

        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            if (matchIndex[row, col] is ExcelEmpty)
                continue;
            var matchKey = Convert.ToInt32(matchIndex[row, col]);

            var matchValue = replaceValue[row, col]?.ToString();

            replacements.Add(matchKey, matchValue);
        }

        var counter = 0;

        var result = Regex.Replace(
            inputRange,
            regexMethod,
            m =>
            {
                counter++;
                if (replacements.TryGetValue(counter, out var expression))
                    return expression;
                return m.Value;
            }
        );

        return result;
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "二维Range查找值，返回指定查找到值的相对行、列"
    )]
    public static object FindValueFromRange(
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "查找值")]
            string seachValue,
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "查找范围")]
            object[,] searchRange,
        [ExcelArgument(
            AllowReference = true,
            Description = "1：行；2：列；其他：自定义行列组{,},[,]",
            Name = "返回值类型"
        )]
            string returnType = "1",
        [ExcelArgument(AllowReference = true, Description = "返回第几个值", Name = "返回值序号")]
            int returnNum = 1
    )
    {
        var rows = searchRange.GetLength(0);
        var cols = searchRange.GetLength(1);

        var counter = 1;
        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            var targetCell = searchRange[row, col];
            if (targetCell is ExcelEmpty)
                continue;

            if (targetCell.ToString() == seachValue)
            {
                if (counter == returnNum)
                {
                    if (returnType == "1")
                        return row + 1;

                    if (returnType == "2")
                        return col + 1;

                    var delimiterList = returnType
                        .ToCharArray()
                        .Select(c => c.ToString())
                        .ToArray();
                    return $"{delimiterList[0]}{row + 1}{delimiterList[1]}{col + 1}{delimiterList[2]}";
                }

                counter++;
            }
        }

        return "不存在";
    }

    [ExcelFunction(
        Category = "UDF-查找值",
        IsVolatile = true,
        IsMacroType = true,
        Description = "判断A列数据是否为0或空，把B列数据累积到向下的非0值位置"
    )]
    public static object SumValueFilterNull(
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "条件范围")]
            object[,] conditionRange,
        [ExcelArgument(AllowReference = true, Description = "单元格地址：A1", Name = "累加范围")]
            object[,] sumRange
    )
    {
        var rows = conditionRange.GetLength(0);
        var cols = conditionRange.GetLength(1);

        var matchesList = new List<double>();

        double sum = 0;

        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            double conditionCell = (double)conditionRange[row, col];

            double sumCell = (double)sumRange[row, col];

            if (conditionCell == 0)
            {
                sum += sumCell;
                matchesList.Add(sum);
            }
            else
            {
                matchesList.Add(sum + sumCell);
                sum = 0;
            }
        }

        return matchesList.ToArray();
    }
}
