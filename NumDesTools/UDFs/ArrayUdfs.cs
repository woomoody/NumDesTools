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
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range，不需要默认值的直接用TEXT JOIN，这个支持默认值，并支持多字符串：首、中、尾拼接"
    )]
    public static string CreatValueToArray(
        [ExcelArgument(
            AllowReference = true,
            Name = "单元格范围",
            Description = "Range&Cell,eg:A1:A2"
        )]
            object[,] rangeObj,
        [ExcelArgument(
            AllowReference = true,
            Name = "默认值单元格范围",
            Description = "Range&Cell,eg:A1:A2，不填表示没有默认值"
        )]
            object[,] rangeObjDef,
        [ExcelArgument(
            AllowReference = true,
            Name = "分隔符",
            Description = "分隔符,默认:[,]表示：首-中-尾符"
        )]
            string delimiter,
        [ExcelArgument(AllowReference = true, Name = "过滤值", Description = "一般为空值或0")]
            string ignoreValue
    )
    {
        if (delimiter == "")
            delimiter = "[,]";
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        string startDelimiter,
            midDelimiter,
            endDelimiter;
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

        var sb = new System.Text.StringBuilder();
        var rows = rangeObj.GetLength(0);
        var cols = rangeObj.GetLength(1);
        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            var item = rangeObj[row, col];
            if (item is ExcelEmpty || item.ToString() == ignoreValue || item is ExcelError)
                continue;
            var val = rangeObjDef[0, 0] is ExcelMissing ? item : rangeObjDef[row, col];
            sb.Append(val).Append(midDelimiter);
        }

        if (sb.Length == 0)
            return Empty;
        sb.Length -= midDelimiter.Length;
        return startDelimiter + sb + endDelimiter;
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range，根据第二个单元格范围内数字重复拼接第一个单元格内对应值"
    )]
    public static string CreatValueToArrayRepeat(
        [ExcelArgument(
            AllowReference = true,
            Name = "单元格范围",
            Description = "Range&Cell,eg:A1:A2"
        )]
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
        var sb = new System.Text.StringBuilder();
        var rows = rangeObj.GetLength(0);
        var cols = rangeObj.GetLength(1);
        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
        {
            var item = rangeObj[row, col];
            if (item is ExcelEmpty || item.ToString() == ignoreValue)
                continue;
            var item2 = rangeObj2[row, col];
#pragma warning disable CA1305
            if (!int.TryParse(item2?.ToString(), out var repeat))
                repeat = 0;
#pragma warning restore CA1305
            for (var i = 0; i < repeat; i++)
                sb.Append(item).Append(delimiter);
        }

        if (sb.Length > 0)
            sb.Length -= delimiter.Length;
        return sb.ToString();
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
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,true/false",
            Name = "过滤空值"
        )]
            bool ignoreEmpty
    )
    {
        var values1Objects = rangeObj1.Cast<object>().ToArray();
        var values2Objects = rangeObj2.Cast<object>().ToArray();
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();

        if (values1Objects.Length == 0 || values2Objects.Length == 0 || delimiterList.Length == 0)
            return Empty;

        var sb = new System.Text.StringBuilder();
        var count = 0;
        foreach (var item in values1Objects)
        {
            var skip =
                ignoreEmpty && (item is ExcelEmpty || string.IsNullOrEmpty(item?.ToString()));
            if (!skip)
            {
                sb.Append(delimiterList[0])
                    .Append(item)
                    .Append(delimiterList[1])
                    .Append(values2Objects[count])
                    .Append(delimiterList[2])
                    .Append(delimiter[1]);
            }
            count++;
        }

        if (sb.Length > 0)
            sb.Length -= 1;
        return delimiterList[0] + sb + delimiterList[2];
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
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg:true/false",
            Name = "过滤空值"
        )]
            string ignoreEmpty,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否包含外围分隔符,eg:,",
            Name = "过滤空值"
        )]
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
            delimiter = "[,]";
        if (ignoreEmpty == "")
            ignoreEmpty = "TRUE";
        if (isOutline == "")
            isOutline = "TRUE";
        // 拼接结果
        var result = Empty;
        // 分隔符处理
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        if (delimiterList.Length < 3)
            throw new ArgumentException("分隔符至少需要三个字符，例如: {,}");

        // 将所有范围转换为二维数组list
        var allValues = new List<object[]>();
        foreach (var range in ranges)
        {
            if (range.ToString() == "ExcelErrorValue")
                continue;
            if (range is object[,] rangeObj)
                allValues.Add(rangeObj.Cast<object>().ToArray());
            else
                throw new ArgumentException("输入的范围必须是二维数组");
        }

        if (allValues.Count == 0)
            return "";

        // 确保所有范围的长度一致
        var maxLength = allValues.Max(arr => arr.Length);
        if (allValues.Any(arr => arr.Length != maxLength))
            throw new ArgumentException("所有单元格范围的长度必须一致");

        for (var i = 0; i < maxLength; i++)
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
                        continue;
                }

                rowValues.Add(value?.ToString() ?? Empty);
            }

            // 拼接每一行的值
            if (rowValues.Count > 0)
                result +=
                    delimiterList[0]
                    + Join(delimiterList[1], rowValues)
                    + delimiterList[2]
                    + delimiter[1];
        }

        // 去掉最后一个多余的分隔符
        if (!IsNullOrEmpty(result))
            result = result.Substring(0, result.Length - 1);

        if (isOutline == "TRUE")
            result = delimiterList[0] + result + delimiterList[2];

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
        [ExcelArgument(
            AllowReference = true,
            Description = "分隔符,eg:[,](头-中-尾)",
            Name = "分隔符"
        )]
            string delimiter,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,true/false",
            Name = "过滤空值"
        )]
            bool ignoreEmpty
    )
    {
        var values1Objects = rangeObj1.Cast<object>().ToArray();
        var values2Objects = rangeObj2.Cast<object>().ToArray();
        var valuesFilterObjects = filterObj.Cast<object>().ToArray();
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();

        if (values1Objects.Length == 0 || values2Objects.Length == 0 || delimiterList.Length == 0)
            return Empty;

        var sb = new System.Text.StringBuilder();
        var count = 0;
        foreach (var item in values1Objects)
        {
            var filterBase = values2Objects[count];
            var matchFilter = ignoreEmpty
                ? filterBase.ToString() == valuesFilterObjects[0].ToString()
                : filterBase == valuesFilterObjects[0];
            var skip = ignoreEmpty && (item is ExcelEmpty || item?.ToString() == "");
            if (!skip && matchFilter)
                sb.Append(item).Append(delimiterList[1]);
            count++;
        }

        if (sb.Length > 0)
            sb.Length -= delimiterList[1].Length;
        return delimiterList[0] + sb + delimiterList[2];
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
            object[,] rangeObj,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj2,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj3,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj4,
        [ExcelArgument(
            AllowReference = true,
            Description = "文本太长会缓存到我的文档",
            Name = "存储文件名"
        )]
            string fileName = "关卡1"
    )
    {
        // 创建一个包含N个数组的对象
        var layers = new object[rangeObj.GetLength(0) * rangeObj.GetLength(1)];
        var layers2 = new object[rangeObj2.GetLength(0) * rangeObj2.GetLength(1)];
        var layers3 = new object[rangeObj3.GetLength(0) * rangeObj3.GetLength(1)];
        var layers4 = new object[rangeObj4.GetLength(0) * rangeObj4.GetLength(1)];

        // 关卡入口数据特殊处理
        var index1List = new List<int> { 22, 23, 30, 31 };
        var index2List = new List<int> { -1, 22, 22, 22 };

        var index = 0;
        for (var i = 0; i < rangeObj.GetLength(0); i++)
        {
            for (var j = 0; j < rangeObj.GetLength(1); j++)
            {
                int[] indexLink = null;
                var linkIndex = -1;
                var range = 0;
                foreach (var te in index1List)
                {
                    if (te == index)
                    {
                        var subIndex = index1List.IndexOf(te);
                        range = 1;
                        if (index2List[subIndex] == -1)
                        {
                            var tempList = new List<int>(index1List); // 创建副本
                            tempList.RemoveAt(0); // 修改副本
                            indexLink = tempList.Select(x => x).ToArray();
                        }
                        else
                        {
                            linkIndex = Convert.ToInt32(index2List[subIndex]);
                        }
                        break;
                    }
                }
                var disrole = 0;
                if (Convert.ToInt32(rangeObj4[i, j]) != -1)
                    disrole = 1;
                layers[index++] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj[i, j]),
                    LinkedIndexes = (object[])null,
                    DisplayRule = 0,
                    LinkedParentIndex = -1,
                    Range = 0,
                    ObstacleConfigId = 0,
                };
                layers2[index - 1] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj2[i, j]),
                    LinkedIndexes = (object[])null,
                    DisplayRule = 0,
                    LinkedParentIndex = -1,
                    Range = 0,
                    ObstacleConfigId = 0,
                };
                layers3[index - 1] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj3[i, j]),
                    LinkedIndexes = indexLink,
                    DisplayRule = 0,
                    LinkedParentIndex = linkIndex,
                    Range = range,
                    ObstacleConfigId = 0,
                };
                layers4[index - 1] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj4[i, j]),
                    LinkedIndexes = indexLink,
                    DisplayRule = disrole,
                    LinkedParentIndex = linkIndex,
                    Range = range,
                    ObstacleConfigId = 0,
                };
            }
        }
        var layersOther = new object[] { layers, layers2, layers3, layers4 };

        var combinedData = new
        {
            Row = rangeObj.GetLength(0),
            Col = rangeObj.GetLength(1),
            GridDataList = (object[])null,
            Layers = layersOther,
            LayerNames = new object[] { "棋子层", "蛛网", "关卡入口", "障碍层" },
        };

        // 将对象转换为 JSON 格式
        var json = JsonConvert.SerializeObject(combinedData, Formatting.None);
        try
        {
            // 获取"我的文档"文件夹路径
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fileNames = $"{fileName}.json";
            string fullPath = Path.Combine(documentsPath, fileNames);

            // 确保目录存在
            Directory.CreateDirectory(
                Path.GetDirectoryName(fullPath) ?? throw new InvalidOperationException()
            );

            // 写入文件
            File.WriteAllText(fullPath, json, Encoding.UTF8);

            PluginLog.Write($"JSON 文件已保存到: {fullPath}");
        }
        catch (Exception ex)
        {
            PluginLog.Write($"保存 JSON 文件失败: {ex.Message}");
        }
        return json;
    }

    [ExcelFunction(
        Category = "UDF-组装字符串",
        IsVolatile = true,
        IsMacroType = true,
        Description = "拼接Range：Range数据转为Json"
    )]
    public static string CreatRangeToJsonV5(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj2,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj3,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObj4,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "第一单元格范围"
        )]
            object[,] rangeObjLock,
        [ExcelArgument(
            AllowReference = true,
            Description = "文本太长会缓存到我的文档",
            Name = "存储文件名"
        )]
            string fileName = "关卡1"
    )
    {
        // 创建一个包含N个数组的对象
        var layers = new object[rangeObj.GetLength(0) * rangeObj.GetLength(1)];
        var layers2 = new object[rangeObj2.GetLength(0) * rangeObj2.GetLength(1)];
        var layers3 = new object[rangeObj3.GetLength(0) * rangeObj3.GetLength(1)];
        var layers4 = new object[rangeObj4.GetLength(0) * rangeObj4.GetLength(1)];

        // 关卡入口数据特殊处理
        var index1List = new List<int>();
        var index2List = new List<int>();
        index2List.Add(-1);
        foreach (var lockItem in rangeObjLock)
        {
            if (lockItem is ExcelEmpty)
            {
                continue;
            }
            int value = Convert.ToInt32(lockItem);
            index1List.Add(value);
        }
        if (index1List.Count == 0)
        {
            index1List = null;
            index2List = null;
        }
        else
        {
            for (int i = 0; i < index1List.Count - 1; i++)
            {
                index2List.Add(index1List[0]);
            }
        }

        var index = 0;
        for (var i = 0; i < rangeObj.GetLength(0); i++)
        {
            for (var j = 0; j < rangeObj.GetLength(1); j++)
            {
                object indexLink = (object[])null;
                var linkIndex = -1;
                var range = 0;

                if (index1List != null)
                {
                    foreach (var te in index1List)
                    {
                        if (te == index)
                        {
                            var subIndex = index1List.IndexOf(te);
                            range = 1;
                            if (index2List[subIndex] == -1)
                            {
                                var tempList = new List<int>(index1List); // 创建副本
                                tempList.RemoveAt(0); // 修改副本
                                indexLink = tempList.Select(x => x).ToArray();
                            }
                            else
                            {
                                linkIndex = Convert.ToInt32(index2List[subIndex]);
                            }
                            break;
                        }
                    }
                }
                var is0layer = 1;
                if (Convert.ToInt32(rangeObj[i, j]) == -1)
                {
                    is0layer = 0;
                }
                var is0layer2 = 1;
                if (Convert.ToInt32(rangeObj2[i, j]) == -1)
                {
                    is0layer2 = 0;
                }
                var is0layer3 = 1;
                if (Convert.ToInt32(rangeObj3[i, j]) == -1)
                {
                    is0layer3 = 0;
                }
                layers[index++] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj[i, j]),
                    LinkedIndexes = (object[])null,
                    DisplayRule = is0layer,
                    LinkedParentIndex = -1,
                    Range = 0,
                    ObstacleConfigId = 0,
                };
                layers2[index - 1] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj2[i, j]),
                    LinkedIndexes = (object[])null,
                    DisplayRule = is0layer2,
                    LinkedParentIndex = -1,
                    Range = 0,
                    ObstacleConfigId = 0,
                };
                layers3[index - 1] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj3[i, j]),
                    LinkedIndexes = (object[])null,
                    DisplayRule = is0layer3,
                    LinkedParentIndex = -1,
                    Range = range,
                    ObstacleConfigId = 0,
                };
                layers4[index - 1] = new
                {
                    Index = index - 1,
                    ConfigId = Convert.ToInt32(rangeObj4[i, j]),
                    LinkedIndexes = indexLink,
                    DisplayRule = 0,
                    LinkedParentIndex = linkIndex,
                    Range = range,
                    ObstacleConfigId = 0,
                };
            }
        }
        var layersOther = new object[] { layers, layers2, layers3, layers4 };

        var combinedData = new
        {
            Row = rangeObj.GetLength(0),
            Col = rangeObj.GetLength(1),
            GridDataList = (object[])null,
            Layers = layersOther,
            LayerNames = new object[] { "棋子层", "蛛网", "障碍物", "特殊障碍物" },
        };

        // 将对象转换为 JSON 格式
        var json = JsonConvert.SerializeObject(combinedData, Formatting.None);
        try
        {
            // 获取"我的文档"文件夹路径
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fileNames = $"{fileName}.json";
            string fullPath = Path.Combine(documentsPath, fileNames);

            // 确保目录存在
            Directory.CreateDirectory(
                Path.GetDirectoryName(fullPath) ?? throw new InvalidOperationException()
            );

            // 写入文件
            File.WriteAllText(fullPath, json, Encoding.UTF8);

            PluginLog.Write($"JSON 文件已保存到: {fullPath}");
        }
        catch (Exception ex)
        {
            PluginLog.Write($"保存 JSON 文件失败: {ex.Message}");
        }
        return json;
    }

    [ExcelFunction(
        Category = "UDF-数组转置",
        IsVolatile = true,
        IsMacroType = true,
        Description = "二维数据转换为一维数据，并可选择是否过滤空值"
    )]
    public static object[,] Trans2ArrayTo1Arrays(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "单元格范围"
        )]
            object[,] rangeObj,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否带索引号",
            Name = "是否带索引号，带2不带1"
        )]
            int outMax,
        [ExcelArgument(
            AllowReference = true,
            Description = "行优先还是列优先：1行0列",
            Name = "行列优先"
        )]
            int rowOrCol,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,true/false",
            Name = "过滤空值"
        )]
            bool ignoreEmpty = true
    )
    {
        //默认值
        if (outMax > 2)
            outMax = 2;

        List<object> rangeValueList = [];
        List<object> rangeColIndexList = [];

        var rowCount = rangeObj.GetLength(0);
        var colCount = rangeObj.GetLength(1);

        if (rowOrCol == 0)
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
        else if (rowOrCol == 1)
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

        var result = new object[rangeValueList.Count, outMax];

        for (var i = 0; i < rangeValueList.Count; i++)
            if (outMax == 2)
            {
                result[i, 1] = rangeValueList[i];
                result[i, 0] = rangeColIndexList[i];
            }
            else
            {
                result[i, 0] = rangeValueList[i];
            }

        return result;
    }

    //[ExcelFunction(
    //    Category = "UDF-数组公式写入",
    //    IsVolatile = true,
    //    IsMacroType = true,
    //    Description = @"针对各种数据公式填充数据慢问题，采用Range写入法",
    //    Name = "RangeWriteFast"
    //)]
    //public static object RangeWriteFast(object[,] inputRange)
    //{
    //    // 获取当前调用位置及目标区域
    //    ExcelReference callerRef = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
    //    int rows = inputRange.GetLength(0);
    //    int cols = inputRange.GetLength(1);

    //    string fullName = (string)XlCall.Excel(XlCall.xlSheetNm, callerRef);
    //    // 匹配格式：[工作簿名]工作表名
    //    Match match = Regex.Match(fullName, @"\]([^!]+)");
    //    string sheetName = match.Success ? match.Groups[1].Value.Trim('\'') : "地编关系";

    //    var app = AppServices.App;
    //    Worksheet targetSheet = app.Sheets[sheetName];

    //    Range targetRange = targetSheet.Range[
    //        targetSheet.Cells[callerRef.RowFirst + 2, callerRef.ColumnFirst + 1],
    //        targetSheet.Cells[callerRef.RowFirst + rows + 1, callerRef.ColumnFirst + cols]
    //    ];

    //    // 直接写入数据
    //    targetRange.Value = inputRange; // 批量写入二维数组

    //    return "最新数据↓";
    //}

    [ExcelFunction(
        Category = "UDF-Excel函数增强",
        IsVolatile = true,
        IsMacroType = true,
        Description = @"针对自动填充，例如：自动填充迭代=IF(C49= "",X48,B49&C49)功能的拓展，输出数组",
        Name = "UXFillBlanks"
    )]
    public static object[,] FillBlanks(object[,] inputRange)
    {
        object[,] result = (object[,])inputRange.Clone();
        object lastValue = null;

        for (int i = 0; i < inputRange.GetLength(0); i++)
        {
            if (inputRange[i, 0] is ExcelEmpty || inputRange[i, 0] == null)
            {
                if (lastValue != null)
                    result[i, 0] = lastValue;
                else
                    result[i, 0] = ExcelError.ExcelErrorNA; // 空值且无前值返回错误
            }
            else
            {
                lastValue = inputRange[i, 0];
            }
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
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "单元格范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "单元格范围"
        )]
            object[,] rangeObj2,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,true/false",
            Name = "过滤空值"
        )]
            bool ignoreEmpty
    )
    {
        List<double> sumProductValueList = [];
        double sumProductValue = 0;

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        for (var col = 0; col < colCount; col++)
        for (var row = 0; row < rowCount; row++)
            if (ignoreEmpty)
            {
                var value1 = rangeObj1[row, col];
                var value2 = rangeObj2[row, col];
                if (double.TryParse(value1.ToString(), out var result1))
                    if (double.TryParse(value2.ToString(), out var result2))
                    {
                        sumProductValue += result1 * result2;
                        sumProductValueList.Add(sumProductValue);
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
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "单元格范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,true/false",
            Name = "过滤空值"
        )]
            bool ignoreEmpty = true
    )
    {
        var uniqueValues = new HashSet<object>();

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        for (var col = 0; col < colCount; col++)
        for (var row = 0; row < rowCount; row++)
        {
            var value = rangeObj1[row, col];

            if (ignoreEmpty && (value == null || IsNullOrEmpty(value.ToString())))
                continue; // 跳过空值

            uniqueValues.Add(value);
        }

        return uniqueValues.Count;
    }

    [ExcelFunction(
        Category = "UDF-Excel函数增强",
        IsVolatile = true,
        IsMacroType = true,
        Description = "指定范围内不重复值的数组",
        Name = "UXUNIQUE"
    )]
    public static object[] UxUnique(
        [ExcelArgument(
            AllowReference = true,
            Description = "Range&Cell,eg:A1:A2",
            Name = "单元格范围"
        )]
            object[,] rangeObj1,
        [ExcelArgument(
            AllowReference = true,
            Description = "是否过滤空值,eg,true/false",
            Name = "过滤空值"
        )]
            bool ignoreEmpty = true
    )
    {
        var uniqueValues = new HashSet<object>();

        var rowCount = rangeObj1.GetLength(0);
        var colCount = rangeObj1.GetLength(1);

        for (var col = 0; col < colCount; col++)
        for (var row = 0; row < rowCount; row++)
        {
            var value = rangeObj1[row, col];
            if (value is ExcelEmpty)
                continue; // 跳过空值
            if (ignoreEmpty && (value == null || IsNullOrEmpty(value.ToString())))
                continue; // 跳过空值

            uniqueValues.Add(value);
        }

        return uniqueValues.ToArray();
    }
}
