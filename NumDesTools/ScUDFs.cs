using System;
using ExcelDna.Integration;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;


namespace NumDesTools;

/// <summary>
/// Excel自定义函数类
/// </summary>
public class ExcelUdf
{
    private static readonly dynamic IndexWk = NumDesAddIn.App.ActiveWorkbook;
    private static readonly dynamic ExcelPath = IndexWk.Path;

    [ExcelFunction(Category = "FindValue", IsVolatile = true, IsMacroType = true, Description = "寻找指定表格字段所在列")]
    public static int FindKeyCol([ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标行")] int row, [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1")
    {
        var path = ExcelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet);
        if (sheet == null) sheet = workbook.GetSheetAt(0);
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

    [ExcelFunction(Category = "FindValue", IsVolatile = true, IsMacroType = true, Description = "寻找指定表格字段所在行")]
    public static int FindKeyRow([ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标列")] int col, [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1")
    {
        var path = ExcelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet);
        if (sheet == null) sheet = workbook.GetSheetAt(0);
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

    [ExcelFunction(Category = "GetExcelInfo", IsVolatile = true, IsMacroType = true, Description = "获取单元格背景色")]
    public static string GetCellColor([ExcelArgument(AllowReference = true, Description = "目标列")] string address)
    {
        var range = NumDesAddIn.App.ActiveSheet.Range[address];
        var color = range.Interior.Color;
        // 将Excel VBA颜色值转换为RGB格式
        var red = (int)(color % 256);
        var green = (int)(color / 256 % 256);
        var blue = (int)(color / 65536 % 256);
        // 返回RGB格式的颜色值
        return $"{red}#{green}#{blue}";
    }

    //拆分字符串
    [ExcelFunction(Category = "StrToNum", IsVolatile = true, IsMacroType = true, Description = "提取字符串中数字")]
    public static int GetNumFromStr([ExcelArgument(AllowReference = true, Description = "输入字符串")] string inputValue,
        [ExcelArgument(AllowReference = true, Description = "分隔符")]
        string delimiter,
        [ExcelArgument(AllowReference = true, Description = "第几个数字")]
        int numCount)
    {
        // 使用正则表达式匹配数字
        var numbers = Regex.Split(inputValue, delimiter)
            .SelectMany(s => Regex.Matches(s, @"\d+").Cast<Match>().Select(m => m.Value))
            .ToArray();
        return Convert.ToInt32(numbers[numCount - 1]);
    }

    //组装字符串
    [ExcelFunction(Category = "StrToNum", IsVolatile = true, IsMacroType = true, Description = "拼接Range")]
    public static string CreatValueToArray(
        [ExcelArgument(AllowReference = true, Description = "单元格范围")]
        object rangeObj,
        [ExcelArgument(AllowReference = true, Description = "默认值单元格范围")]
        object rangeObjDef,
        [ExcelArgument(AllowReference = true, Description = "分隔符")]
        string delimiter,
        [ExcelArgument(AllowReference = true, Description = "过滤值")]
        string ignoreValue,
        [ExcelArgument(AllowReference = true, Description = "返回值类型")]
        int returnType)
    {
        // 将传递的 object 类型参数转换为 Range 对象
        var rangeRef = (ExcelReference)rangeObj;
        var rangeRefDef = (ExcelReference)rangeObjDef;
        // 使用 ExcelReference.GetValue 获取选定范围的值
        var values = (object[,])rangeRef.GetValue();
        var valuesDef = (object[,])rangeRefDef.GetValue();
        //过滤掉空值并将二维数组中的值按行拼接成字符串
        var result = string.Empty;
        var count = 0;
        foreach (var item in values)
        {
            if (item is ExcelEmpty || item.ToString() == ignoreValue)
            {
            }
            else
            {
                if (returnType != 0)
                {
                    var itemDef = valuesDef[0, count];
                    result += itemDef + delimiter;
                }
                else
                {
                    result += item + delimiter;
                }
            }

            count++;
        }

        if (result != "") result = result.Substring(0, result.Length - 1);
        return result;
    }

    //组装字符串(二维)
    [ExcelFunction(Category = "StrToNum", IsVolatile = true, IsMacroType = true, Description = "拼接Range（二维）")]
    public static string CreatValueToArray2(
        [ExcelArgument(AllowReference = true, Description = "单元格范围1")]
        object rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "单元格范围2")]
        object rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "分隔符")]
        string delimiter)
    {
        // 将传递的 object 类型参数转换为 Range 对象
        var rangeRef1 = (ExcelReference)rangeObj1;
        var rangeRef2 = (ExcelReference)rangeObj2;
        // 使用 ExcelReference.GetValue 获取选定范围的值
        var values1 = (object[,])rangeRef1.GetValue();
        var values2 = (object[,])rangeRef2.GetValue();
        //变为一维数组
        var values1Objects = values1.Cast<object>().ToArray();
        var values2Objects = values2.Cast<object>().ToArray();
        //获取间隔方案
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        //过滤掉空值并将二维数组中的值按行拼接成字符串
        var result = string.Empty;
        var count = 0;
        foreach (var item in values1Objects)
        {
            var itemDef = delimiterList[0] + item + delimiterList[1] + values2Objects[count] + delimiterList[2];
            result += itemDef + delimiter[1];
            count++;
        }

        result = delimiterList[0] + result + delimiterList[2];
        return result;
    }
}