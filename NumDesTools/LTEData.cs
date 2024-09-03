using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using NPOI.OpenXmlFormats.Dml.Diagram;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;

namespace NumDesTools;

public class LteData
{
    private static readonly dynamic Wk = NumDesAddIn.App.ActiveWorkbook;

    private static readonly string WkPath = Wk.Path;

    //导出LTE数据配置
    public static void ExportLteDataConfig(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sw = new Stopwatch();
        sw.Start();

        //获取【导出】表信息
        var ws = Wk.ActiveSheet;
        var selectRange = NumDesAddIn.App.Selection;
        var baseSheetName = selectRange.Value2.ToString();
        var selectRow = selectRange.Row;
        var selectCol = selectRange.Column;

        //基本信息
        var exportBaseSheetData = new Dictionary<string, Dictionary<string, Tuple<int, int>>>();
        var exportBaseData = new Dictionary<string, Tuple<int, int>>();
        if (exportBaseData == null)
            throw new ArgumentNullException(nameof(exportBaseData));

        var baseRangeValue = ws.Range[
            ws.Cells[selectRow, selectCol + 2],
            ws.Cells[selectRow + 2, selectCol + 11]
        ].Value2;

        for (int col = 1; col <= 10; col++)
        {
            var keyName = baseRangeValue[1, col]?.ToString() ?? "";
            if (keyName != "")
            {
                var keyCol = (int)baseRangeValue[2, col];
                var keyRowMax = (int)baseRangeValue[3, col];
                exportBaseData[keyName] = new Tuple<int, int>(keyCol, keyRowMax);
            }
        }
        exportBaseSheetData[baseSheetName] = exportBaseData;

        //通配符信息
        var exportWildcardData = new Dictionary<string, string>();
        if (exportWildcardData == null)
            throw new ArgumentNullException(nameof(exportWildcardData));

        var wildcardCount = (int)ws.Cells[selectRow + 1, selectCol].Value2;
        var wildcardRangeValue = ws.Range[
            ws.Cells[selectRow, selectCol + 13],
            ws.Cells[selectRow + wildcardCount - 1, selectCol + 14]
        ].Value2;
        for (int row = 1; row <= wildcardCount; row++)
        {
            var wildcardName = wildcardRangeValue[row, 1]?.ToString() ?? "";
            if (wildcardName != "")
            {
                var wildcardValue = wildcardRangeValue[row, 2]?.ToString() ?? "";
                exportWildcardData[wildcardName] = wildcardValue;
            }
        }

        //读取【基础/任务……】表数据
        var baseSheet = Wk.Worksheets[baseSheetName];
        var baseData = new Dictionary<string, List<object>>();
        var baseSheetData = exportBaseSheetData[baseSheetName];

        foreach (var baseElement in baseSheetData)
        {
            var range = baseSheet
                .Range[
                    baseSheet.Cells[2, baseElement.Value.Item1],
                    baseSheet.Cells[baseElement.Value.Item2, baseElement.Value.Item1]
                ]
                .Value2;

            var dataList = PubMetToExcel.List2DToListRowOrCol(
                PubMetToExcel.RangeDataToList(range),
                true
            );

            baseData[baseElement.Key] = dataList;
        }

        //获取【#LTE数据模版】信息
        var modelSheet = Wk.Worksheets["#LTE数据模版"];
        var modelListObjects = modelSheet.ListObjects;
        var modelValueAll = new Dictionary<string, Dictionary<(object, object), string>>();

        foreach (ListObject list in modelListObjects)
        {
            var modelName = list.Name;
            var modelRangeValue = list.Range.Value2;

            int rowCount = modelRangeValue.GetLength(0);
            int colCount = modelRangeValue.GetLength(1);

            // 将二维数组的数据存储到字典中
            var modelValue = PubMetToExcel.Array2DToDic2D(rowCount, colCount, modelRangeValue);
            if (modelValue == null)
            {
                return;
            }
            modelValueAll[modelName] = modelValue;
        }
        //分功能处理导出
        if (baseSheetName.Contains("【基础】"))
        {
            var idList = baseData["ID"];
            var nameList = baseData["当前包装"];
            var typeList = baseData["类型"];
            var linkMaxList = baseData["链长"];
            var areaList = baseData["首次出现"];

            //走【基础】表逻辑
            LteBaseSheet(idList, nameList, typeList, linkMaxList , areaList , exportWildcardData, modelValueAll) ;
        }
        else if (baseSheetName.Contains("【任务】"))
        {
            //走【任务】表逻辑
        }

        sw.Stop();
        var ts2 = sw.Elapsed;
        NumDesAddIn.App.StatusBar = "导出完成，用时：" + ts2;
    }

    private static void LteBaseSheet(
        List<object> idList,
        List<object> nameList,
        List<object> typeList,
        List<object> linkMaxList,
        List<object> areaList,
        Dictionary<string, string> exportWildcardData,
        Dictionary<string, Dictionary<(object, object), string>> modelValueAll
    )
    {
        Dictionary<string, Dictionary<(object, object), string>> realValueAll;
        var strDictionary = new Dictionary<string, Dictionary<string, List<string>>>();

        //替换通配符生成数据
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        List<(string, string, string)> errorList = PubMetToExcel.SetExcelObjectEpPlus(
            WkPath,
            "Item.xlsx",
            out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel
        );

        for (int i = 0; i < idList.Count; i++)
        {
            string itemIndex = idList[i]?.ToString() ?? "";
            if (itemIndex == "")
                continue;

            string itemType = typeList[i].ToString();
            string itemName = nameList[i].ToString();
            int itemLinkMax = (int)linkMaxList[i];

            var defaultItem = new List<string> { itemIndex, itemName, itemType };

            var writeRow = targetSheet.Dimension.End.Row;
            var writeCol = targetSheet.Dimension.End.Column;
            for (int j = 2; j <= writeCol; j++)
            {
                var cellTitle = targetSheet.Cells[2, j].Value?.ToString() ?? "";
                if (cellTitle == "")
                    continue;
                // 使用 LINQ 查询判断字典中是否包含指定的值
                bool containsValue = modelValueAll["Item.xlsx"]
                    .Keys.Any(key => key.Item1.Equals(itemType) && key.Item2.Equals(cellTitle));

                if (containsValue)
                {
                    var cellModelValue = modelValueAll["Item.xlsx"][(itemType, cellTitle)];
                    //分析cellModelValue中的通配符
                    var cellRealValue = AnalyzeWildcard(
                        cellModelValue,
                        exportWildcardData,
                        defaultItem,
                        strDictionary
                    );

                    var cell = targetSheet.Cells[writeRow + 1, j];
                    cell.Value = cellRealValue;
                }
            }
        }
        targetExcel.Save();
        targetSheet.Dispose();
    }

    private static void TaskSheet(string specialCharsStr) { }

    //分析Cell中通配符构成
    private static string AnalyzeWildcard(
        string cellModelValue,
        Dictionary<string, string> exportWildcardData,
        List<string> defaultItem,
        Dictionary<string, Dictionary<string, List<string>>> strDictionary
    )
    {
        string fixWildcardValue = "";
        string cellRealValue = cellModelValue;
        string pattern = "#(.*?)#";

        MatchCollection matches = Regex.Matches(cellModelValue, pattern);

        foreach (Match match in matches)
        {
            var wildcard = match.Groups[1].Value;
            switch (wildcard)
            {
                //动态值基础上计算值
                case "物品详细类型":
                    fixWildcardValue = CutString(exportWildcardData, "物品详细类型");
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "物品名称编号":
                    fixWildcardValue = SetString(exportWildcardData, "物品名称编号");
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "物品图鉴组编号":
                    var valueDic = SetStringDic(exportWildcardData, "物品图鉴组编号");
                    fixWildcardValue = valueDic.Item1;
                    var fixWildcardValue2 = valueDic.Item2;
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    //初始化字典
                    InitializeDictionary(strDictionary, "物品图鉴组编号", fixWildcardValue);
                    //添加值
                    strDictionary["物品图鉴组编号"][fixWildcardValue].Add(fixWildcardValue2);
                    break;
                case "合成结果编号":
                    fixWildcardValue = MerString(exportWildcardData, "合成结果编号");
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "合成结果编号（蛛网）":
                    fixWildcardValue = MerSpiString(exportWildcardData, "合成结果编号（蛛网）");
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "合成返还编号":
                    fixWildcardValue = MerReString(exportWildcardData, "合成返还编号");
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "建造转换编号":
                    fixWildcardValue = BuildString(exportWildcardData, "建造转换编号");
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "物品吸附":
                    fixWildcardValue = AdsString(exportWildcardData, "物品吸附");
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;

                //获取动态值
                case "物品编号":
                    fixWildcardValue = GetString(exportWildcardData, "物品编号", defaultItem);
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "物品备注":
                    fixWildcardValue = GetString(exportWildcardData, "物品名称", defaultItem);
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "物品类型":
                    fixWildcardValue = GetString(exportWildcardData, "物品类型", defaultItem);
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "链类最大值":
                    fixWildcardValue = GetString(exportWildcardData, "链类最大值", defaultItem);
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;
                case "物品区域":
                    fixWildcardValue = GetString(exportWildcardData, "物品区域", defaultItem);
                    cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
                    break;

                //获取静态默认值
                default:
                    cellRealValue = cellRealValue.Replace(
                        $"#{wildcard}#",
                        exportWildcardData[wildcard]
                    );
                    break;
            }
        }
        return cellRealValue;
    }

    //自定义字典初始化
    private static void InitializeDictionary(
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        string outerKey,
        string innerKey
    )
    {
        if (!strDictionary.ContainsKey(outerKey))
        {
            strDictionary[outerKey] = new Dictionary<string, List<string>>();
        }

        if (!strDictionary[outerKey].ContainsKey(innerKey))
        {
            strDictionary[outerKey][innerKey] = new List<string>();
        }
    }

    //截取字符串
    private static string CutString(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        return ProcessString(
            exportWildcardData,
            targetType,
            (result, dynamicParam1, _) => result.Substring(0, dynamicParam1),
            8
        );
    }

    //重置字符串
    private static string SetString(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        return ProcessString(
            exportWildcardData,
            targetType,
            (result, dynamicParam1, _) => result.Substring(0, result.Length - dynamicParam1) + "00",
            2
        );
    }

    //重置字符串（元组化）
    private static Tuple<string, string> SetStringDic(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        var result = SetString(exportWildcardData, targetType);
        return new Tuple<string, string>(result, exportWildcardData[targetType]);
    }
    
    //合成字符串
    private static string MerString(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        return ProcessString(
            exportWildcardData,
            targetType,
            (result, dynamicParam1, _) => (long.Parse(result) + dynamicParam1).ToString(),
            1
        );
    }

    //合成蛛网字符串
    private static string MerSpiString(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        return ProcessString(
            exportWildcardData,
            targetType,
            (result, dynamicParam1, dynamicParam2) =>
                (long.Parse(result) + dynamicParam1 - dynamicParam2).ToString(),
            1,
            30
        );
    }

    //合成返回字符串
    private static string MerReString(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        return ProcessString(
            exportWildcardData,
            targetType,
            (result, dynamicParam1, _) => (long.Parse(result) - dynamicParam1).ToString(),
            30
        );
    }

    //建造字符串
    private static string BuildString(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        return ProcessString(
            exportWildcardData,
            targetType,
            (result, dynamicParam1, _) => (long.Parse(result) - dynamicParam1).ToString(),
            10
        );
    }

    //建造字符串
    private static string AdsString(
        Dictionary<string, string> exportWildcardData,
        string targetType
    )
    {
        return ProcessString(
            exportWildcardData,
            targetType,
            (result, dynamicParam1, _) => (long.Parse(result) - dynamicParam1).ToString(),
            10
        );
    }

    //获取动态字符串
    private static string GetString(
        Dictionary<string, string> exportWildcardData,
        string targetType,
        List<string> defaultIndex
    )
    {
        if (targetType == "物品编号")
        {
            exportWildcardData[targetType] = defaultIndex[0];
        }
        else if (targetType == "物品名称")
        {
            exportWildcardData[targetType] = defaultIndex[1];
        }
        else if (targetType == "物品类型")
        {
            exportWildcardData[targetType] = defaultIndex[2];
        }
        else if (targetType == "链类最大值")
        {
            exportWildcardData[targetType] = defaultIndex[3];
        }
        else if (targetType == "物品区域")
        {
            exportWildcardData[targetType] = defaultIndex[4];
        }
        return exportWildcardData[targetType];
    }
    //字符串处理策略（fuc改成liststring，存储物品编号，链类最大值等信息？用来根据物品变成生成数组字符串
    private static string ProcessString(
        Dictionary<string, string> exportWildcardData,
        string targetType,
        Func<string, int, int, string> processFunc,
        int dynamicParam1,
        int dynamicParam2 = 0
    )
    {
        var pattern = exportWildcardData[targetType];
        var splitResult = Regex.Split(pattern, "#");
        int param1 = splitResult.Length > 2 ? int.Parse(splitResult[2]) : dynamicParam1; //两个#后的参数
        int param2 = splitResult.Length > 3 ? int.Parse(splitResult[3]) : dynamicParam2;
        var result = exportWildcardData[splitResult[1]];
        return processFunc(result, param1, param2);
    }
}
