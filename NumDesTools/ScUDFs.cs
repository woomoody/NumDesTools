using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

#pragma warning disable CA1416


namespace NumDesTools;

/// <summary>
/// Excel自定义函数类
/// </summary>
public class ExcelUdf
{
    private static readonly dynamic IndexWk = NumDesAddIn.App.ActiveWorkbook;
    private static readonly dynamic ExcelPath = IndexWk.Path;

    [ExcelFunction(Category = "UDF-查找值", IsVolatile = true, IsMacroType = true, Description = "寻找同层级指定表格字段所在列")]
    public static int FindKeyCol([ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标行")] int row, [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1")
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

    [ExcelFunction(Category = "UDF-查找值", IsVolatile = true, IsMacroType = true, Description = "寻找同层级指定表格字段所在行")]
    public static int FindKeyRow([ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标列")] int col,
        [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1")
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
    [ExcelFunction(Category = "UDF-查找值", IsVolatile = true, IsMacroType = true, Description = "寻找同层级指定表格字段所在列指定行的值")]
    public static string FindKeyColToRow([ExcelArgument(Description = "工作簿")] string targetWorkbook,
    [ExcelArgument(Description = "目标行")] int row, [ExcelArgument(Description = "输出目标行")] int rowOut, [ExcelArgument(Description = "匹配值")] string searchValue,
    [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1" )
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

    [ExcelFunction(Category = "UDF-查找值", IsVolatile = true, IsMacroType = true, Description = "寻找同层级指定表格字段所在行指定列的值")]
    public static string FindKeyRowToCol([ExcelArgument(Description = "工作簿")] string targetWorkbook,
        [ExcelArgument(Description = "目标列")] int col, [ExcelArgument(Description = "输出目标列")] int outCol,
        [ExcelArgument(Description = "匹配值")] string searchValue,
        [ExcelArgument(Description = "工作表")] string targetSheet = "Sheet1")
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
    [ExcelFunction(Category = "UDF-查找值", IsVolatile = true, IsMacroType = true, Description = "兼容索引，索引单元格有值则相对索引，否则绝对索引，索引最靠近的单元格（上-左）")]
    public static string FindKeyClose([ExcelArgument(AllowReference = true, Description = "单元格地址：A1" , Name = "单元格")] object inputRange,
        [ExcelArgument(Description = "行索引或列索引")] bool isRow)
    {
        if (inputRange is ExcelReference cellRef)
        {
            var sheet = NumDesAddIn.App.ActiveSheet;
            var rangeRow = cellRef.RowFirst + 1;
            var rangeCol = cellRef.ColumnFirst + 1;
            var rangeValue = sheet.Cells[rangeRow, rangeCol].Value;
            if (rangeValue== null)
            {
                if (isRow)
                {
                    int count = rangeRow;
                    while (count > 0)
                    {
                        var newRangeValue = sheet.Cells[count, rangeCol].Value;
                        if (newRangeValue == null)
                        {
                            count--;
                        }
                        else
                        {
                            return newRangeValue.ToString();
                        }
                    }
                }
                else
                {
                    int count = rangeCol;
                    while (count > 0)
                    {
                        var newRangeValue = sheet.Cells[rangeRow, count].Value;
                        if (newRangeValue == null)
                        {
                            count--;
                        }
                        else
                        {
                            return newRangeValue.ToString();
                        }
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
    [ExcelFunction(Category = "UDF-获取表格信息", IsVolatile = true, IsMacroType = true, Description = "获取单元格背景色")]
    public static string GetCellColor([ExcelArgument(AllowReference = true, Name = "单元格地址",Description = "引用Range&Cell地址,eg:A1")] object address)
    {
        if (address is ExcelReference cellRef)
        {
            var sheet = NumDesAddIn.App.ActiveSheet;
            var rangeRow = cellRef.RowFirst + 1;
            var rangeCol = cellRef.ColumnFirst + 1;
            var range = sheet.Cells[rangeRow, rangeCol];
            var color = range.Interior.Color;
            // 将Excel VBA颜色值转换为RGB格式
            var red = (int)(color % 256);
            var green = (int)(color / 256 % 256);
            var blue = (int)(color / 65536 % 256);
            // 返回RGB格式的颜色值
            return $"{red}#{green}#{blue}";
        }
        return "error";
    }
    [ExcelFunction(Category = "UDF-设置表格信息", IsVolatile = true, IsMacroType = true, Description = "设置单元格背景色")]
    public static string SetCellColor([ExcelArgument(AllowReference = true, Name = "单元格值", Description = "获取单元格值")] string inputValue)
    {
        //使用该公式的单元格地址
        ExcelReference cellRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
        string address = (string)XlCall.Excel(XlCall.xlfReftext, cellRef, true);
        var sheet = NumDesAddIn.App.ActiveSheet;
        var range = sheet.Range[address];
        bool canConvertToInt = int.TryParse(inputValue, out int intValue);
        if (!canConvertToInt)
        {
            return "error";
        }
        var value = intValue % 2;
        if (value == 0)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Aquamarine);
        }
        else
        {
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BurlyWood);
        }
        return "^0^";
    }
    //拆分字符串为int数字
    [ExcelFunction(Category = "UDF-字符串提取数字", IsVolatile = true, IsMacroType = true, Description = "提取字符串中数字")]
    public static int GetNumFromStr([ExcelArgument(AllowReference = true, Description = "输入字符串")] string inputValue,
        [ExcelArgument(AllowReference = true, Name = "分隔符",Description = "分隔符,eg:,")]
        string delimiter,
        [ExcelArgument(AllowReference = true, Name = "数字序号",Description = "选择提取字符串中的第几个数字，如果值很大，表示提取最末尾字符")]
        int numCount)
    {
        // 使用正则表达式匹配数字
        var numbers = Regex.Split(inputValue, delimiter)
            .SelectMany(s => Regex.Matches(s, @"\d+").Select(m => m.Value))
            .ToArray();
        //增加只提取末尾字符的判断
        var maxNumCount = numbers.Length;
        numCount = Math.Min(maxNumCount, numCount);
#pragma warning disable CA1305 // 指定 IFormatProvider
        return Convert.ToInt32(numbers[numCount - 1]);
#pragma warning restore CA1305 // 指定 IFormatProvider
    }
    //拆分字符串为Str字符串
    [ExcelFunction(Category = "UDF-字符串提取数字", IsVolatile = true, IsMacroType = true, Description = "分割字符串为若干字符串")]
    public static string GetStrFromStr([ExcelArgument(AllowReference = true, Name="单元格索引",Description = "输入字符串")] string inputValue,
        [ExcelArgument(AllowReference = true, Name = "分隔符",Description = "分隔符,eg:,")]
        string delimiter,
        [ExcelArgument(AllowReference = true, Name = "过滤符",Description = "过滤符,eg:[,]")]
        string filter,
        [ExcelArgument(AllowReference = true, Name = "序号",Description = "选择提取字符串中的第几个字符串，如果值很大，表示提取最末尾字符")]
        int numCount)
    {
        // 分割字符串
        var filterGroup = filter.ToCharArray().Select(c => c.ToString()).ToArray();
        var strGroup = Regex.Split(inputValue, delimiter);
        if (filterGroup.Length > 0)
        {
            foreach (var filterItem in filterGroup)
            {
                for (int i =0;i< strGroup.Length;i++)
                {
                    strGroup[i]= strGroup[i].Replace(filterItem, "");
                }
            }
        }
        //增加只提取末尾字符的判断
        var maxNumCount = strGroup.Length;
        numCount = Math.Min(maxNumCount, numCount);
        //返回
        return strGroup[numCount - 1];
    }
    //组装字符串
    [ExcelFunction(Category = "UDF-组装字符串", IsVolatile = true, IsMacroType = true, Description = "拼接Range，不需要默认值的直接用TEXT JOIN，这个支持默认值")]
    public static string CreatValueToArray(
        [ExcelArgument(AllowReference = true, Name = "单元格范围" ,Description ="Range&Cell,eg:A1:A2")]
        object[,] rangeObj,
        [ExcelArgument(AllowReference = true, Name = "默认值单元格范围",Description ="Range&Cell,eg:A1:A2")]
        object[,] rangeObjDef,
        [ExcelArgument(AllowReference = true, Name = "分隔符",Description ="分隔符,eg:,")]
        string delimiter,
        [ExcelArgument(AllowReference = true, Name = "过滤值",Description ="一般为空值")]
        string ignoreValue,
        [ExcelArgument(AllowReference = true, Name = "返回值类型",Description ="0指使用默认值模式，非0为一般模式")]
        int returnType)
    {
        //过滤掉空值并将二维数组中的值按行拼接成字符串
        var result = string.Empty;
        var rows = rangeObj.GetLength(0);
        var cols = rangeObj.GetLength(1);
        for (int row = 0; row < rows; row++)
        {
            for (int col = 0; col < cols; col++)
            {
                var item = rangeObj[row, col];
                if (item is ExcelEmpty || item.ToString() == ignoreValue)
                {
                }
                else
                {
                    if (returnType != 0)
                    {
                        var itemDef = rangeObjDef[row, col];
                        result += itemDef + delimiter;
                    }
                    else
                    {
                        result += item + delimiter;
                    }
                }
            }
        }

        if (result != "") result = result.Substring(0, result.Length - 1);
        return result;
    }
    //组装字符串，按数字重复填写ID
    [ExcelFunction(Category = "UDF-组装字符串", IsVolatile = true, IsMacroType = true, Description = "拼接Range，根据第二个单元格范围内数字重复拼接第一个单元格内对应值")]
    public static string CreatValueToArrayRepeat(
        [ExcelArgument(AllowReference = true, Name = "单元格范围" ,Description ="Range&Cell,eg:A1:A2")]
        object[,] rangeObj,
        [ExcelArgument(AllowReference = true, Name = "单元格范围-数量" ,Description ="Range&Cell,eg:A1:A2")]
        object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Name = "分隔符",Description ="分隔符,eg:,")]
        string delimiter,
        [ExcelArgument(AllowReference = true, Name = "过滤值",Description ="一般为空值")]
        string ignoreValue)
    {
        //过滤掉空值并将二维数组中的值按行拼接成字符串
        var result = string.Empty;
        var rows = rangeObj.GetLength(0);
        var cols = rangeObj.GetLength(1);
        for (int row = 0; row < rows; row++)
        {
            for (int col = 0; col < cols; col++)
            {
                var item = rangeObj[row, col];
                if (item is ExcelEmpty || item.ToString() == ignoreValue)
                {
                }
                else
                {
                    var item2 = rangeObj2[row, col];
#pragma warning disable CA1305 // 指定 IFormatProvider
                    for (int i = 0; i < Convert.ToInt32(item2); i++)
                    {
                        result += item + delimiter;
                    }
#pragma warning restore CA1305 // 指定 IFormatProvider
                }
            }
        }
        if (result != "") result = result.Substring(0, result.Length - 1);
        return result;
    }

    //组装字符串(二维)
    [ExcelFunction(Category = "UDF-组装字符串", IsVolatile = true, IsMacroType = true, Description = "拼接Range（二维）")]
    public static string CreatValueToArray2(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "第一单元格范围")]
        object[,] rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2",Name = "第二单元格范围")]
        object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "分隔符,eg:,",Name = "分隔符")]
        string delimiter, 
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false",Name = "过滤空值")]
        bool ignoreEmpty
        )

    {
        //变为一维数组
        var values1Objects = rangeObj1.Cast<object>().ToArray();
        var values2Objects = rangeObj2.Cast<object>().ToArray();
        //获取间隔方案
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        //过滤掉空值并将二维数组中的值按行拼接成字符串
        var result = string.Empty;
        var count = 0;
        if(values1Objects.Length > 0 && values2Objects.Length > 0 && delimiterList.Length >0 )
        {
            foreach (var item in values1Objects)
            {
                if (ignoreEmpty)
                {
                    var excelNull = item is ExcelEmpty;
                    var stringNull = ReferenceEquals(item.ToString(), "");
                    if ( !excelNull && !stringNull )
                    {
                        var itemDef = delimiterList[0] + item + delimiterList[1] + values2Objects[count] + delimiterList[2];
                        result += itemDef + delimiter[1];
                        count++;
                    }
                }
                else
                {
                    var itemDef = delimiterList[0] + item + delimiterList[1] + values2Objects[count] + delimiterList[2];
                    result += itemDef + delimiter[1];
                    count++;
                }
            }
            result = result.Substring(0, result.Length - 1);
            result = delimiterList[0] + result+ delimiterList[2];

        }
        return result;
    }
    //组装字符串：条件
    [ExcelFunction(Category = "UDF-组装字符串", IsVolatile = true, IsMacroType = true, Description = "拼接Range：条件")]
    public static string CreatValueToArrayFilter(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "第一单元格范围")]
        object[,] rangeObj1,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2",Name = "第二单元格范围")]
        object[,] rangeObj2,
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1",Name = "第二个单元格筛选条件值")]
        object[,] filterObj,
        [ExcelArgument(AllowReference = true, Description = "分隔符,eg:[,](头-中-尾)",Name = "分隔符")]
        string delimiter,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false",Name = "过滤空值")]
        bool ignoreEmpty
        )

    {
        //变为一维数组
        var values1Objects = rangeObj1.Cast<object>().ToArray();
        var values2Objects = rangeObj2.Cast<object>().ToArray();
        var valuesFilterObjects = filterObj.Cast<object>().ToArray();

        //获取间隔方案
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        //过滤掉空值并将二维数组中的值按行拼接成字符串
        var result = string.Empty;
        var count = 0;
        if (values1Objects.Length > 0 && values2Objects.Length > 0 && delimiterList.Length > 0)
        {
            foreach (object item in values1Objects)
            {
                if (ignoreEmpty)
                {
                    var excelNull = item is ExcelEmpty;
                    var stringNull = item?.ToString();
                    if (!excelNull && stringNull != "")
                    {
                        var filterObjectBase = values2Objects[count];
                        if (filterObjectBase.ToString() == valuesFilterObjects[0].ToString())
                        {
                             result += item + delimiterList[1];
                        }
                    }
                    count++;
                }
                else
                {
                    var filterObjectBase = values2Objects[count];
                    if (filterObjectBase == valuesFilterObjects[0])
                    {
                        result += item + delimiterList[1];
                    }
                    count++;
                }
            }
            if (!string.IsNullOrEmpty(result))
            {
                result = result.Substring(0, result.Length - 1);
            }
            result = delimiterList[0] + result + delimiterList[2];

        }
        return result;
    }
    //转置二维数组为一维
    [ExcelFunction(Category = "UDF-数组转置", IsVolatile = true, IsMacroType = true, Description = "二维数据转换为一维数据，并可选择是否过滤空值")]
    public static object[,] Trans2ArrayTo1Arrays(
        [ExcelArgument(AllowReference = true, Description = "Range&Cell,eg:A1:A2", Name = "单元格范围")] object[,] rangeObj,
        [ExcelArgument(AllowReference = true, Description = "是否过滤空值,eg,true/false",Name = "过滤空值")] bool ignoreEmpty)
    {
        List<object> rangeValueList = [];
        List<object> rangeColIndexList = [];

        int rowCount = rangeObj.GetLength(0);
        int colCount = rangeObj.GetLength(1);

        for (int col = 0; col < colCount; col++)
        {
            for (int row = 0; row < rowCount; row++)
            {
                object value = rangeObj[row , col ];

                if (ignoreEmpty)
                {
                    var excelNull = value is ExcelEmpty;
                    var stringNull = ReferenceEquals(value.ToString(), "");
                    if (!excelNull && !stringNull)
                    {
                        rangeValueList.Add(value);
                        rangeColIndexList.Add(col +1);
                    }
                }
                else
                {
                    rangeValueList.Add(value);
                    rangeColIndexList.Add(col+1 );
                }
            }
        }
        object[,] result = new object[rangeValueList.Count,2];

        for (int i = 0; i < rangeValueList.Count; i++)
        {
            result[i, 1] = rangeValueList[i];
            result[i, 0] = rangeColIndexList[i];
        }
        return result;
    }
}