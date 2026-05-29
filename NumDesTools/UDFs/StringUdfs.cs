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
    [ExcelFunction(
        Category = "UDF-获取表格信息",
        IsVolatile = true,
        IsMacroType = true,
        Description = "获取单元格背景色"
    )]
    public static string GetCellColor(
        [ExcelArgument(
            AllowReference = true,
            Name = "单元格地址",
            Description = "引用Range&Cell地址,eg:A1"
        )]
            object address
    )
    {
        if (address is ExcelReference cellRef)
        {
            var sheet = AppServices.App.ActiveSheet;
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
        var sheet = AppServices.App.ActiveSheet;
        var range = sheet.Range[address];
        var canConvertToInt = int.TryParse(inputValue, out var intValue);
        if (!canConvertToInt)
            return "error";
        var value = intValue % 2;

        int colorCode =
            value == 0
                ? 0x7FFFD4 // Aquamarine 的 RGB 值
                : 0xDEB887; // BurlyWood 的 RGB 值

        range.Interior.Color = colorCode;
        return "^0^";
    }

    [ExcelFunction(
        Category = "UDF-字符串提取数字",
        IsVolatile = true,
        IsMacroType = true,
        Description = "提取字符串中数字"
    )]
    public static object GetNumFromStr(
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

        if (maxNumCount >= numCount)
        {
            return Convert.ToInt64(numbers[numCount - 1]);
        }
        else
        {
            return "";
        }
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
        [ExcelArgument(
            AllowReference = true,
            Name = "正则方法",
            Description = @"正则方法,eg:\[(.*?)\]"
        )]
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
            regexStr = @"\[(\d+.*?)]";
        // 正则表达式匹配内部数组
        var matches = Regex.Matches(inputValue, regexStr);
        if (numCount > matches.Count)
            numCount = matches.Count;
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
            delimiter = ",";

        var matchesList = new List<string>();
        foreach (var value in inputValue)
        {
            // 正则表达式匹配内部数组
            var numbers = Regex
                .Split(value.ToString() ?? throw new InvalidOperationException(), delimiter)
                .SelectMany(s => Regex.Matches(s, @"\d+").Select(m => m.Value))
                .ToArray();
            foreach (var num in numbers)
                matchesList.Add(num);
        }

        return matchesList.ToArray();
    }
}
