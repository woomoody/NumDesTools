using System.Text.RegularExpressions;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;
using static System.String;
using Match = System.Text.RegularExpressions.Match;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// Excel自定义函数类
/// </summary>
public class ExcelUdf
{
    private static readonly dynamic IndexWk = NumDesAddIn.App.ActiveWorkbook;
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
                    var outColSource = sheet.GetRow(outCol);
                    var outCell = outColSource.GetCell(i);
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
            var sheet = NumDesAddIn.App.ActiveSheet;
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
        string result = Regex.Replace(inputRange, regexMethod, replaceValue);

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
        {
            return "输入单元格地址不能为空";
        }

        if (IsNullOrEmpty(regexMethod))
        {
            regexMethod = @"\d+";
        }

        try
        {
            // 使用正则表达式匹配
            var matches = Regex.Matches(inputRange, regexMethod);

            // 将匹配结果连接为字符串
            string result = Join(", ", matches.Select(m => m.Value));

            // 如果没有匹配到内容，返回提示信息
            if (IsNullOrEmpty(result))
            {
                return "未找到匹配内容";
            }

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

        Dictionary<int, string> replacements = new Dictionary<int, string>();

        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            if (matchIndex[row, col] is ExcelEmpty)
            {
                continue;
            }
            int matchKey = Convert.ToInt32(matchIndex[row, col]);

            string matchValue = replaceValue[row, col]?.ToString();

            replacements.Add(matchKey, matchValue);
        }

        int counter = 0;

        var result = Regex.Replace(
            inputRange,
            regexMethod,
            m =>
            {
                counter++;
                if (replacements.TryGetValue(counter, out var expression))
                {
                    return expression;
                }
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

        int counter = 1;
        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            var targetCell = searchRange[row, col];
            if (targetCell is ExcelEmpty)
            {
                continue;
            }

            if (targetCell.ToString() == seachValue)
            {
                if (counter == returnNum)
                {
                    if (returnType == "1")
                    {
                        return row + 1;
                    }

                    if (returnType == "2")
                    {
                        return col + 1;
                    }

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
        Category = "UDF-获取表格信息",
        IsVolatile = true,
        IsMacroType = true,
        Description = "获取单元格背景色"
    )]
    public static string GetCellColor(
        [ExcelArgument(AllowReference = true, Name = "单元格地址", Description = "引用Range&Cell地址,eg:A1")]
            object address
    )
    {
        if (address is ExcelReference cellRef)
        {
            var sheet = NumDesAddIn.App.ActiveSheet;
            var rangeRow = cellRef.RowFirst + 1;
            var rangeCol = cellRef.ColumnFirst + 1;
            var range = sheet.Cells[rangeRow, rangeCol];
            var color = range.Interior.Color;
            var red = (int)(color % 256);
            var green = (int)(color / 256 % 256);
            var blue = (int)(color / 65536 % 256);
            return $"{red}#{green}#{blue}";
        }

        return "error";
    }

    [ExcelFunction(
        Category = "UDF-设置表格信息",
        IsVolatile = true,
        IsMacroType = true,
        Description = "设置单元格背景色"
    )]
    public static string SetCellColor(
        [ExcelArgument(AllowReference = true, Name = "单元格值", Description = "获取单元格值")]
            string inputValue
    )
    {
        var cellRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        var address = (string)XlCall.Excel(XlCall.xlfReftext, cellRef, true);
        var sheet = NumDesAddIn.App.ActiveSheet;
        var range = sheet.Range[address];
        var canConvertToInt = int.TryParse(inputValue, out var intValue);
        if (!canConvertToInt)
            return "error";
        var value = intValue % 2;
        range.Interior.Color = ColorTranslator.ToOle(
            value == 0 ? Color.Aquamarine : Color.BurlyWood
        );
        return "^0^";
    }

    [ExcelFunction(
        Category = "UDF-字符串提取数字",
        IsVolatile = true,
        IsMacroType = true,
        Description = "提取字符串中数字"
    )]
    public static long GetNumFromStr(
        [ExcelArgument(AllowReference = true, Description = "输入字符串")] string inputValue,
        [ExcelArgument(AllowReference = true, Name = "分隔符", Description = "分隔符,eg:,")]
            string delimiter,
        [ExcelArgument(
            AllowReference = true,
            Name = "数字序号",
            Description = "选择提取字符串中的第几个数字，如果值很大，表示提取最末尾字符"
        )]
            int numCount
    )
    {
        var numbers = Regex
            .Split(inputValue, delimiter)
            .SelectMany(s => Regex.Matches(s, @"\d+").Select(m => m.Value))
            .ToArray();
        var maxNumCount = numbers.Length;
        numCount = Math.Min(maxNumCount, numCount);
#pragma warning disable CA1305
        return Convert.ToInt64(numbers[numCount - 1]);
#pragma warning restore CA1305
    }

    [ExcelFunction(
        Category = "UDF-字符串提取数字",
        IsVolatile = true,
        IsMacroType = true,
        Description = "分割字符串为若干字符串"
    )]
    public static string GetStrFromStr(
        [ExcelArgument(AllowReference = true, Name = "单元格索引", Description = "输入字符串")]
            string inputValue,
        [ExcelArgument(AllowReference = true, Name = "分隔符", Description = "分隔符,eg:,")]
            string delimiter,
        [ExcelArgument(AllowReference = true, Name = "过滤符", Description = "过滤符,eg:[,]")]
            string filter,
        [ExcelArgument(
            AllowReference = true,
            Name = "序号",
            Description = "选择提取字符串中的第几个字符串，如果值很大，表示提取最末尾字符"
        )]
            int numCount
    )
    {
        var filterGroup = filter.ToCharArray().Select(c => c.ToString()).ToArray();
        var strGroup = Regex.Split(inputValue, delimiter);
        if (filterGroup.Length > 0)
            foreach (var filterItem in filterGroup)
                for (var i = 0; i < strGroup.Length; i++)
                    strGroup[i] = strGroup[i].Replace(filterItem, "");
        var maxNumCount = strGroup.Length;
        numCount = Math.Min(maxNumCount, numCount);
        return strGroup[numCount - 1];
    }

    [ExcelFunction(
        Category = "UDF-字符串提取数字",
        IsVolatile = true,
        IsMacroType = true,
        Description = "分割字符串为特定结构的若干字符串"
    )]
    public static string GetStrStructFromStr(
        [ExcelArgument(AllowReference = true, Name = "单元格索引", Description = "输入字符串")]
            string inputValue,
        [ExcelArgument(AllowReference = true, Name = "正则方法", Description = @"正则方法,eg:\[(.*?)\]")]
            string regexStr,
        [ExcelArgument(
            AllowReference = true,
            Name = "序号",
            Description = "选择提取字符串中的第几个字符串，如果值很大，表示提取最末尾字符"
        )]
            int numCount
    )
    {
        if (regexStr == "")
        {
            regexStr = @"\[(\d+.*?)]";
        }
        // 正则表达式匹配内部数组
        var matches = Regex.Matches(inputValue, regexStr);
        if (numCount > matches.Count)
        {
            numCount = matches.Count;
        }
        // 使用 ElementAtOrDefault 安全地获取匹配项
        var match = matches.ElementAtOrDefault(numCount - 1);
        // 如果匹配项存在，则返回其值，否则返回空字符串
        return match?.Value ?? Empty;
    }

    [ExcelFunction(
        Category = "UDF-字符串提取数字",
        IsVolatile = true,
        IsMacroType = true,
        Description = "分割字符串为特定结构的若干字符串-返回数组"
    )]
    public static object GetStrStructFromStrArray(
        [ExcelArgument(AllowReference = true, Name = "单元格索引", Description = "输入字符串")]
            object[,] inputValue,
        [ExcelArgument(AllowReference = true, Name = "分割符", Description = @"默认为逗号")]
            string delimiter
    )
    {
        if (delimiter == "")
        {
            delimiter = ",";
        }

        var matchesList = new List<string>();
        foreach (var value in inputValue)
        {
            // 正则表达式匹配内部数组
            var numbers = Regex
                .Split(value.ToString() ?? throw new InvalidOperationException(), delimiter)
                .SelectMany(s => Regex.Matches(s, @"\d+").Select(m => m.Value))
                .ToArray();
            foreach (var num in numbers)
            {
                matchesList.Add(num);
            }
        }

        return matchesList.ToArray();
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range，不需要默认值的直接用TEXT JOIN，这个支持默认值，并支持多字符串：首、中、尾拼接"
    )]
    public static string CreatValueToArray(
        [ExcelArgument(AllowReference = true, Name = "单元格范围", Description = "Range&Cell,eg:A1:A2")]
            object[,] rangeObj,
        [ExcelArgument(
            AllowReference = true,
            Name = "默认值单元格范围",
            Description = "Range&Cell,eg:A1:A2，不填表示没有默认值"
        )]
            object[,] rangeObjDef,
        [ExcelArgument(AllowReference = true, Name = "分隔符", Description = "分隔符,默认:[,]表示：首-中-尾符")]
            string delimiter,
        [ExcelArgument(AllowReference = true, Name = "过滤值", Description = "一般为空值")]
            string ignoreValue
    )
    {
        var result = Empty;
        //设定默认值
        if (delimiter == "")
        {
            delimiter = "[,]";
        }
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        string startDelimiter;
        string midDelimiter;
        string endDelimiter;
        if (delimiterList.Length == 3)
        {
            startDelimiter = delimiterList[0];
            midDelimiter = delimiterList[1];
            endDelimiter = delimiterList[2];
        }
        else
        {
            startDelimiter = Empty;
            midDelimiter = delimiterList[0];
            endDelimiter = Empty;
        }
        var rows = rangeObj.GetLength(0);
        var cols = rangeObj.GetLength(1);
        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            var item = rangeObj[row, col];
            if (item is ExcelEmpty || item.ToString() == ignoreValue || item is ExcelError) { }
            else
            {
                if (!(rangeObjDef[0, 0] is ExcelMissing))
                {
                    var itemDef = rangeObjDef[row, col];
                    result += itemDef + midDelimiter;
                }
                else
                {
                    result += item + midDelimiter;
                }
            }
        }

        if (result != "")
            result = startDelimiter + result.Substring(0, result.Length - 1) + endDelimiter;
        return result;
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range，根据第二个单元格范围内数字重复拼接第一个单元格内对应值"
    )]
    public static string CreatValueToArrayRepeat(
        [ExcelArgument(AllowReference = true, Name = "单元格范围", Description = "Range&Cell,eg:A1:A2")]
            object[,] rangeObj,
        [ExcelArgument(
            AllowReference = true,
            Name = "单元格范围-数量",
            Description = "Range&Cell,eg:A1:A2"
        )]
            object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Name = "分隔符", Description = "分隔符,eg:,")]
            string delimiter,
        [ExcelArgument(AllowReference = true, Name = "过滤值", Description = "一般为空值")]
            string ignoreValue
    )
    {
        var result = Empty;
        var rows = rangeObj.GetLength(0);
        var cols = rangeObj.GetLength(1);
        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            var item = rangeObj[row, col];
            if (item is ExcelEmpty || item.ToString() == ignoreValue) { }
            else
            {
                var item2 = rangeObj2[row, col];
#pragma warning disable CA1305
                for (var i = 0; i < Convert.ToInt32(item2); i++)
                    result += item + delimiter;
#pragma warning restore CA1305
            }
        }

        if (result != "")
            result = result.Substring(0, result.Length - 1);
        return result;
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range（二维）"
    )]
    public static string CreatValueToArray2(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第二单元格范围"
        )]
            object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "分隔符,eg:,", Name = "分隔符")]
            string delimiter,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false", Name = "过滤空值")]
            bool ignoreEmpty
    )
    {
        var values1Objects = rangeObj1.Cast<object>().ToArray();
        var values2Objects = rangeObj2.Cast<object>().ToArray();
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        var result = Empty;

        if (values1Objects.Length > 0 && values2Objects.Length > 0 && delimiterList.Length > 0)
        {
            var count = 0;
            foreach (var item in values1Objects)
            {
                if (ignoreEmpty)
                {
                    var excelNull = item is ExcelEmpty;
                    var stringNull = ReferenceEquals(item.ToString(), "");
                    if (!excelNull && !stringNull && item.ToString() != "")
                    {
                        var itemDef =
                            delimiterList[0]
                            + item
                            + delimiterList[1]
                            + values2Objects[count]
                            + delimiterList[2];
                        result += itemDef + delimiter[1];
                    }
                }
                else
                {
                    var itemDef =
                        delimiterList[0]
                        + item
                        + delimiterList[1]
                        + values2Objects[count]
                        + delimiterList[2];
                    result += itemDef + delimiter[1];
                }
                count++;
            }

            result = result.Substring(0, result.Length - 1);
            result = delimiterList[0] + result + delimiterList[2];
        }

        return result;
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range（二维）-动态参数",
        ExplicitRegistration = true
    )]
    public static string CreatValueToArray2Dya(
        [ExcelArgument(AllowReference = true, Description = "分隔符,eg:,", Name = "分隔符")]
            string delimiter,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg:true/false", Name = "过滤空值")]
            string ignoreEmpty,
        [ExcelArgument(AllowReference = true, Description = "是否包含外围分隔符,eg:,", Name = "过滤空值")]
            string isOutline,
        [ExcelArgument(
            AllowReference = false,
            Description = "多个单元格范围，支持动态输入，eg: A1:A2, B1:B2",
            Name = "单元格范围"
        )]
            params object[] ranges
    )
    {
        //默认值
        if (delimiter == "")
        {
            delimiter = "[,]";
        }
        if (ignoreEmpty == "")
        {
            ignoreEmpty = "TRUE";
        }
        if (isOutline == "")
        {
            isOutline = "TRUE";
        }
        // 拼接结果
        var result = Empty;
        // 分隔符处理
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        if (delimiterList.Length < 3)
        {
            throw new ArgumentException("分隔符至少需要三个字符，例如: {,}");
        }

        // 将所有范围转换为二维数组list
        var allValues = new List<object[]>();
        foreach (var range in ranges)
        {
            if (range.ToString() == "ExcelErrorValue")
                continue;
            if (range is object[,] rangeObj)
            {
                allValues.Add(rangeObj.Cast<object>().ToArray());
            }
            else
            {
                throw new ArgumentException("输入的范围必须是二维数组");
            }
        }

        if (allValues.Count == 0)
        {
            return "";
        }

        // 确保所有范围的长度一致
        var maxLength = allValues.Max(arr => arr.Length);
        if (allValues.Any(arr => arr.Length != maxLength))
        {
            throw new ArgumentException("所有单元格范围的长度必须一致");
        }

        for (int i = 0; i < maxLength; i++)
        {
            var rowValues = new List<string>();
            foreach (var rangeValues in allValues)
            {
                var value = rangeValues[i];
                if (ignoreEmpty == "TRUE")
                {
                    var isExcelEmpty = value is ExcelEmpty;
                    var isExcelError = value is ExcelError;
                    var isStringEmpty = value?.ToString() == Empty;
                    if (isExcelEmpty || isStringEmpty || isExcelError)
                    {
                        continue;
                    }
                }
                rowValues.Add(value?.ToString() ?? Empty);
            }

            // 拼接每一行的值
            if (rowValues.Count > 0)
            {
                result +=
                    delimiterList[0]
                    + Join(delimiterList[1], rowValues)
                    + delimiterList[2]
                    + delimiter[1];
            }
        }

        // 去掉最后一个多余的分隔符
        if (!IsNullOrEmpty(result))
        {
            result = result.Substring(0, result.Length - 1);
        }

        if (isOutline == "TRUE")
        {
            result = delimiterList[0] + result + delimiterList[2];
        }

        return result;
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range：条件"
    )]
    public static string CreatValueToArrayFilter(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第二单元格范围"
        )]
            object[,] rangeObj2,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1",
            Name = "第二个单元格筛选条件值"
        )]
            object[,] filterObj,
        [ExcelArgument(AllowReference = true, Description = "分隔符,eg:[,](头-中-尾)", Name = "分隔符")]
            string delimiter,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false", Name = "过滤空值")]
            bool ignoreEmpty
    )
    {
        var values1Objects = rangeObj1.Cast<object>().ToArray();
        var values2Objects = rangeObj2.Cast<object>().ToArray();
        var valuesFilterObjects = filterObj.Cast<object>().ToArray();

        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        var result = Empty;
        var count = 0;
        if (values1Objects.Length > 0 && values2Objects.Length > 0 && delimiterList.Length > 0)
        {
            foreach (var item in values1Objects)
                if (ignoreEmpty)
                {
                    var excelNull = item is ExcelEmpty;
                    var stringNull = item?.ToString();
                    if (!excelNull && stringNull != "")
                    {
                        var filterObjectBase = values2Objects[count];
                        if (filterObjectBase.ToString() == valuesFilterObjects[0].ToString())
                            result += item + delimiterList[1];
                    }

                    count++;
                }
                else
                {
                    var filterObjectBase = values2Objects[count];
                    if (filterObjectBase == valuesFilterObjects[0])
                        result += item + delimiterList[1];
                    count++;
                }

            if (!IsNullOrEmpty(result))
                result = result.Substring(0, result.Length - 1);
            result = delimiterList[0] + result + delimiterList[2];
        }

        return result;
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range：Range数据转为Json"
    )]
    public static string CreatRangeToJson(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第二单元格范围"
        )]
            object[,] rangeObj2
    )
    {
        // 创建一个包含两个数组的对象
        var gridDataList = new object[rangeObj1.GetLength(0) * rangeObj1.GetLength(1)];
        int index = 0;
        for (int i = 0; i < rangeObj1.GetLength(0); i++)
        {
            for (int j = 0; j < rangeObj1.GetLength(1); j++)
            {
                gridDataList[index++] = new
                {
                    ConfigId = Convert.ToInt32(rangeObj1[i, j]),
                    ObstacleConfigId = Convert.ToInt32(rangeObj2[i, j])
                };
            }
        }

        var combinedData = new
        {
            GridDataList = gridDataList,
            Row = rangeObj1.GetLength(0),
            Col = rangeObj2.GetLength(1)
        };

        // 将对象转换为 JSON 格式
        string json = JsonConvert.SerializeObject(combinedData, Formatting.None);

        return json;
    }

    [ExcelFunction(
        Category = "UDF-数组转置",
        IsVolatile = true,
        IsMacroType = true,
        Description = "二维数据转换为一维数据，并可选择是否过滤空值"
    )]
    public static object[,] Trans2ArrayTo1Arrays(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "单元格范围")]
            object[,] rangeObj,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false", Name = "过滤空值")]
            bool ignoreEmpty,
        [ExcelArgument(AllowReference = true, Description = "行优先还是列优先：0行1列", Name = "行列优先")]
            int rowOrCol
    )
    {
        List<object> rangeValueList = [];
        List<object> rangeColIndexList = [];

        var rowCount = rangeObj.GetLength(0);
        var colCount = rangeObj.GetLength(1);

        if (rowOrCol == 0)
        {
            //按行
            for (var col = 0; col < colCount; col++)
            for (var row = 0; row < rowCount; row++)
            {
                var value = rangeObj[row, col];

                if (ignoreEmpty)
                {
                    var excelNull = value is ExcelEmpty;
                    var stringNull = ReferenceEquals(value.ToString(), "");
                    if (!excelNull && !stringNull)
                    {
                        rangeValueList.Add(value);
                        rangeColIndexList.Add(col + 1);
                    }
                }
                else
                {
                    rangeValueList.Add(value);
                    rangeColIndexList.Add(col + 1);
                }
            }
        }
        else if (rowOrCol == 1)
        {
            //按行
            for (var row = 0; row < rowCount; row++)
            for (var col = 0; col < colCount; col++)
            {
                var value = rangeObj[row, col];

                if (ignoreEmpty)
                {
                    var excelNull = value is ExcelEmpty;
                    var stringNull = ReferenceEquals(value.ToString(), "");
                    if (!excelNull && !stringNull)
                    {
                        rangeValueList.Add(value);
                        rangeColIndexList.Add(col + 1);
                    }
                }
                else
                {
                    rangeValueList.Add(value);
                    rangeColIndexList.Add(col + 1);
                }
            }
        }

        var result = new object[rangeValueList.Count, 2];

        for (var i = 0; i < rangeValueList.Count; i++)
        {
            result[i, 1] = rangeValueList[i];
            result[i, 0] = rangeColIndexList[i];
        }

        return result;
    }

    [ExcelFunction(
        Category = "UDF-Excel函数增强",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对原生Excel函数SUMPRODUCT功能的拓展，输出数组",
        Name = "UXSUMPRODUCT"
    )]
    public static double[] UxSumProduct(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "单元格范围")]
            object[,] rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "单元格范围")]
            object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false", Name = "过滤空值")]
            bool ignoreEmpty
    )
    {
        List<double> sumProductValueList = [];
        double sumProductValue = 0;

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        for (var col = 0; col < colCount; col++)
        for (var row = 0; row < rowCount; row++)
        {
            if (ignoreEmpty)
            {
                var value1 = rangeObj1[row, col];
                var value2 = rangeObj2[row, col];
                if (double.TryParse(value1.ToString(), out double result1))
                {
                    if (double.TryParse(value2.ToString(), out double result2))
                    {
                        sumProductValue += result1 * result2;
                        sumProductValueList.Add(sumProductValue);
                    }
                }
            }
        }

        return sumProductValueList.ToArray();
    }
    [ExcelFunction(
        Category = "UDF-Excel函数增强",
        IsVolatile = true,
        IsMacroType = true,
        Description = "统计指定范围内不重复值的数量",
        Name = "UXUNIQUECOUNT"
    )]
    public static int UxUniqueCount(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "单元格范围")]
        object[,] rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false", Name = "过滤空值")]
        bool ignoreEmpty = true
    )
    {
        HashSet<object> uniqueValues = new HashSet<object>();

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        for (var col = 0; col < colCount; col++)
        {
            for (var row = 0; row < rowCount; row++)
            {
                var value = rangeObj1[row, col];

                if (ignoreEmpty && (value == null || IsNullOrEmpty(value.ToString())))
                {
                    continue; // 跳过空值
                }

                uniqueValues.Add(value);
            }
        }

        return uniqueValues.Count;
    }
    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数",
        Name = "AliceCountBuildLinks"
    )]
    public static double[] AliceCountBuildLinks(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "链等级范围")]
            object[,] rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "链数量范围")]
            object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "矿数量范围")]
            object[,] rangeObj3,
        [ExcelArgument(AllowReference = true, Description = "链最大级别,默认为：8", Name = "数量范围")]
            int linkMax,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,0/其他数组，不填为0", Name = "过滤空值")]
            int ignoreEmpty
    )
    {
        if (linkMax == 0)
        {
            linkMax = 8;
        }
        double[] linkTotalCount = new double[linkMax];

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        var buildRowCount = rangeObj3.GetLength(0);
        var buildColCount = rangeObj3.GetLength(1);

        for (var col = 0; col < colCount; col++)
        {
            object buildNumResult = null;
            if (buildRowCount == 1 && col < buildColCount)
            {
                buildNumResult = rangeObj3[0, col];
            }
            else if (buildColCount == 1 && col < buildRowCount)
            {
                buildNumResult = rangeObj3[col, 0];
            }

            if (
                buildNumResult != null
                && double.TryParse(buildNumResult.ToString(), out double buildNum)
            )
            {
                for (var row = 0; row < rowCount; row++)
                {
                    if (ignoreEmpty == 0)
                    {
                        var linkLevelResult = rangeObj1[row, col];
                        var linkNumResult = rangeObj2[row, col];
                        if (
                            linkLevelResult != null
                            && linkNumResult != null
                            && int.TryParse(linkLevelResult.ToString(), out int linkLevel)
                            && double.TryParse(linkNumResult.ToString(), out double linkNum)
                        )
                        {
                            if (linkLevel > 0 && linkLevel <= linkMax)
                            {
                                linkTotalCount[linkLevel - 1] += linkNum * buildNum;
                            }
                        }
                    }
                }
            }
        }
        return linkTotalCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数",
        Name = "AliceCountBuilds"
    )]
    public static double[] AliceCountBuilds(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "矿数量范围")]
            object[,] rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "矿最大级别,默认为：3", Name = "矿最大等级")]
            int buildMax,
        [ExcelArgument(AllowReference = true, Description = "数据按列还是按行：默认按列：0，1为按行", Name = "按行按列")]
            int rowOrCol,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,0/其他数组，不填为0", Name = "过滤空值")]
            int ignoreEmpty
    )
    {
        if (buildMax == 0)
        {
            buildMax = 3;
        }
        double[] buildTotalCount = new double[buildMax];

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        bool processRows = rowOrCol == 0;
        int outerLoopCount = processRows ? rowCount : colCount;
        //int innerLoopCount = processRows ? colCount : rowCount;

        for (var outer = 0; outer < outerLoopCount; outer++)
        {
            object buildLevelResult = processRows ? rangeObj1[outer, 0] : rangeObj1[0, outer];
            object buildNumResult = processRows ? rangeObj1[outer, 1] : rangeObj1[1, outer];

            if (ignoreEmpty == 0 && buildLevelResult != null && buildNumResult != null)
            {
                if (
                    int.TryParse(buildLevelResult.ToString(), out int buildLevel)
                    && double.TryParse(buildNumResult.ToString(), out double buildNum)
                )
                {
                    if (buildLevel > 0 && buildLevel <= buildMax)
                    {
                        buildTotalCount[buildLevel - 1] += buildNum;
                    }
                }
            }
        }

        return buildTotalCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数",
        Name = "AliceCountMergeLinks"
    )]
    public static double AliceCountMergeLinks(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "链数量范围")]
            object[,] rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "链积分范围")]
            object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "链最大级别,默认为：8", Name = "链最大等级")]
            int linksMax,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,0/其他数组，不填为0", Name = "过滤空值")]
            int ignoreEmpty
    )
    {
        if (linksMax == 0)
        {
            linksMax = 8;
        }

        double mergeScoreTotalCount = 0;

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        int addLinksNum = 0;

        for (var row = 0; row < rowCount - 1; row++)
        {
            for (var col = 0; col < colCount; col++)
            {
                object linksNumResult = rangeObj1[row, col];
                object linksScoreResult = rangeObj2[row, col];

                if (ignoreEmpty == 0 && linksNumResult != null && linksScoreResult != null)
                {
                    if (
                        int.TryParse(linksNumResult.ToString(), out int linksNum)
                        && double.TryParse(linksScoreResult.ToString(), out double linksScore)
                    )
                    {
                        linksNum += addLinksNum;
                        addLinksNum = (int)(linksNum / 2.5);
                        //倒数第2链或者大于等于5时才合成的积分
                        if (row >= linksMax - 3 || linksNum >= 5)
                        {
                            mergeScoreTotalCount += addLinksNum * linksScore;
                        }
                    }
                }
            }
        }

        return mergeScoreTotalCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-计算合成N个M阶链消耗",
        Name = "AliceCountLinksMax"
    )]
    public static double AliceCountLinksMax(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "链最大等级")]
            int rangeObjMax,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "链最大等级需要数量")]
            int rangeObjMaxCount,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "合成类型")]
            double mergeType,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "链1已有数量")]
            double rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "链2已有数量")]
            double rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "链3已有数量")]
            double rangeObj3,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "链4已有数量")]
            double rangeObj4
    )
    {
        double baseLinkCount = rangeObjMaxCount;
        var hasLinks = new List<double> { rangeObj4, rangeObj3, rangeObj2, rangeObj1 };

        for (var row = 1; row < rangeObjMax; row++)
        {
            if (row >= rangeObjMax - 3)
            {
                baseLinkCount =
                    Math.Ceiling(baseLinkCount * mergeType) - hasLinks[row - (rangeObjMax - 4)];
            }
            else
            {
                baseLinkCount = Math.Ceiling(baseLinkCount * mergeType);
            }
        }

        return baseLinkCount;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-计算最近坐标",
        Name = "AliceLtePoisonNear"
    )]
    public static string AliceLtePoisonNear(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "基准坐标")]
            object[,] basePos,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "目标坐标组")]
            string targetPos,
        [ExcelArgument(AllowReference = true, Description = "\\d+", Name = "目标坐标组正则方法")]
            string posPattern,
        [ExcelArgument(AllowReference = true, Description = "1/0/-1", Name = "选择：最近、最远、中值")]
            string posType
    )
    {
        var baseX = int.Parse(basePos[0, 0].ToString() ?? throw new InvalidOperationException());
        var baseY = int.Parse(basePos[0, 1].ToString() ?? throw new InvalidOperationException());
        posPattern ??= "";
        posType ??= "1";
        // 提取坐标
        MatchCollection posMatches = Regex.Matches(targetPos, posPattern);
        // 构建结果
        var posResult = new List<(double, string)>();
        if (posResult == null)
            throw new ArgumentNullException(nameof(posResult));
        foreach (Match match in posMatches)
        {
            int x = int.Parse(match.Groups[1].Value);
            int y = int.Parse(match.Groups[2].Value);
            double distance = Math.Pow(x - baseX, 2) + Math.Pow(y - baseY, 2);
            posResult.Add((distance, $"{x},{y}"));
        }

        //选择结果
        if (posType == "1")
        {
            var minValueTuple = posResult.MinBy(t => t.Item1);
            return minValueTuple.Item2;
        }

        if (posType == "0")
        {
            var maxValueTuple = posResult.MaxBy(t => t.Item1);
            return maxValueTuple.Item2;
        }

        var sortedList = posResult.OrderBy(t => t.Item1).ToList();
        var middleIndex = sortedList.Count / 2;
        var medianValueTuple = sortedList[middleIndex];
        return medianValueTuple.Item2;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-提取坐标",
        Name = "AliceLtePoison"
    )]
    public static string AliceLtePoison(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "目标坐标组")]
            string targetPos,
        [ExcelArgument(AllowReference = true, Description = "\\d+", Name = "目标坐标组正则方法")]
            string posPattern
    )
    {
        // 提取坐标
        MatchCollection posMatches = Regex.Matches(targetPos, posPattern);
        // 构建结果
        string posResult = Empty;
        if (posResult == null)
            throw new ArgumentNullException(nameof(posResult));
        foreach (Match match in posMatches)
        {
            int x = int.Parse(match.Groups[1].Value);
            int y = int.Parse(match.Groups[2].Value);
            var pos = "{" + $"21,{x},{y}" + "},";
            posResult += pos;
        }
        posResult = posResult.Substring(0, posResult.Length - 1);
        return posResult;
    }

    [ExcelFunction(
        Category = "UDF-Alice专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "针对Alice项目特制的自定义函数-查找指定Range文件名数据在文件夹中是否存在",
        Name = "AliceLteSourceCheck"
    )]
    public static bool AliceLteSourceCheck(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "文件名")]
            string filesName,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "目标文件夹子目录 ")]
            string folderPath,
        [ExcelArgument(AllowReference = true, Description = @"C:\My", Name = "目标文件夹根目录")]
            string baseFolderPath
    )
    {
        baseFolderPath = IsNullOrEmpty(baseFolderPath) ? @"C:/M1Work/Code/" : baseFolderPath;

        var fileFullPath = baseFolderPath + folderPath;
        try
        {
            // 获取文件夹及其子文件夹中的所有文件
            var files = Directory.GetFiles(fileFullPath, filesName, SearchOption.AllDirectories);
            if (files.Length > 0)
            {
                return true;
            }
        }
        catch (Exception)
        {
            return false;
        }

        return false;
    }

    [ExcelFunction(
        Category = "UDF-ChatGPT专属函数",
        IsVolatile = true,
        IsMacroType = true,
        Description = "使用ChatGPT辅助翻译-反应还是比较慢",
        Name = "ChatTransfer"
    )]
    public static object ChatTransfer(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "要翻译的单元格")]
            object[,] sourceLan,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1", Name = "要翻译的语言类型")]
            object[,] lanType,
        [ExcelArgument(
            AllowReference = true,
            Description = "补充对翻译的要求，例如：语言，格式：默认：英语",
            Name = "翻译要求"
        )]
            string addContent,
        [ExcelArgument(AllowReference = true, Description = "任意字符串或缺省", Name = "忽略空值")]
            string ignoreValue
    )
    {
        // 使用 ExcelAsyncUtil.Run 实现异步操作
        return ExcelAsyncUtil.Run(
            "ChatTransfer",
            new object[] { sourceLan, lanType, addContent, ignoreValue },
            () =>
            {
                try
                {
                    // 获取 API Key、Url 、model
                    var apiKey = NumDesAddIn.ApiKey;
                    var apiUrl = NumDesAddIn.ApiUrl;
                    var apiModel = NumDesAddIn.ApiModel;

                    // 处理 sourceLan 数据
                    var sourceLanStr = ProcessInputRange(sourceLan, ignoreValue, @"\n");

                    // 处理 lanType 数据
                    var lanTypeStr = ProcessInputRange(lanType, ignoreValue, ",");

                    // 构造系统提示内容
                    var sysContent = NumDesAddIn.ChatGptSysContentTransferAss + "翻译为：" + lanTypeStr;

                    // 构造请求体
                    object requestBody = null;
                    if (apiModel.Contains("gpt"))
                    {
                        requestBody = new
                        {
                            model = apiModel,
                            messages = new[]
                            {
                                new { role = "system", content = sysContent },
                                new { role = "user", content = sourceLanStr }
                            },
                            max_tokens = 10000
                        };
                    }
                    else if (apiModel.Contains("deepseek"))
                    {
                        requestBody = new
                        {
                            model = apiModel,

                            messages = new[]
                            {
                                new { content = sysContent, role = "system" },
                                new { content = sourceLanStr, role = "user" }
                            },
                            max_tokens = 2048,
                            frequency_penalty = 0,
                            presence_penalty = 0,
                            response_format = new { type = "text" },
                            stop = (string)null,
                            stream = false,
                            stream_options = (object)null,
                            temperature = 1,
                            top_p = 1,
                            tools = (object)null,
                            tool_choice = "none",
                            logprobs = false,
                            top_logprobs = (object)null
                        };
                    }

                    // 调用 Chat API
                    string response = ChatApiClient
                        .CallApiAsync(requestBody, apiKey, apiUrl)
                        .GetAwaiter()
                        .GetResult();

                    // 解析返回结果
                    var responseList = JsonConvert.DeserializeObject<List<List<object>>>(response);

                    // 使用 LINQ 转置
                    responseList = responseList[0]
                        .Select((_, colIndex) => responseList.Select(row => row[colIndex]).ToList())
                        .ToList();

                    var responseRange = PubMetToExcel.ConvertListToArray(responseList);

                    return responseRange;
                }
                catch (Exception ex)
                {
                    // 捕获异常并返回错误信息
                    return $"Error: {ex.Message}";
                }
            }
        );
    }

    // 处理输入范围数据
    private static string ProcessInputRange(
        object[,] inputRange,
        string ignoreValue,
        string delimiter
    )
    {
        var result = new List<string>();

        foreach (var item in inputRange)
        {
            if (item is ExcelEmpty || item is ExcelError || item.ToString() == ignoreValue)
            {
                continue;
            }
            result.Add(item.ToString());
        }

        return Join(delimiter, result);
    }
}
