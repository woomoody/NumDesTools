using System.Collections.Generic;
using System.Reflection.Metadata.Ecma335;
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
            ws.Cells[selectRow + wildcardCount, selectCol + 14]
        ].Value2;
        for (int row = 1; row <= wildcardCount; row++)
        {
            var wildcardName = wildcardRangeValue[row, 1]?.ToString() ?? "";
            if (wildcardName != "")
            {
                var wildcardValue = wildcardRangeValue[row, 2].ToString();
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
            //走【基础】表逻辑
            BaseSheet(baseData, exportWildcardData, modelValueAll);
        }
        else if (baseSheetName.Contains("【任务】"))
        {
            //走【任务】表逻辑
        }

        sw.Stop();
        var ts2 = sw.Elapsed;
        NumDesAddIn.App.StatusBar = "导出完成，用时：" + ts2;
    }

    private static void BaseSheet(
        Dictionary<string, List<object>> baseData,
        Dictionary<string, string> exportWildcardData,
        Dictionary<string, Dictionary<(object, object), string>> modelValueAll
    )
    {
        Dictionary<string, Dictionary<(object, object), string>> realValueAll;
        var strDictionary = new Dictionary<string, Dictionary<string, List<string>>>();

        //替换通配符生成数据
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        foreach (var modelSheet in modelValueAll)
        {
            string modelSheetName = modelSheet.Key;

            List<(string, string, string)> errorList = PubMetToExcel.SetExcelObjectEpPlus(
                WkPath,
                modelSheetName,
                out ExcelWorksheet targetSheet,
                out ExcelPackage targetExcel
            );
            var idList = baseData["ID"];
            var typeList = baseData["类型"];

            var writeCol = targetSheet.Dimension.End.Column;

            var exportWildcardDyData = new Dictionary<string, string>(exportWildcardData);

            for (int idCount = 0; idCount < idList.Count; idCount++)
            {
                string itemId = idList[idCount]?.ToString() ?? "";
                if (itemId == "")
                    continue;
                string itemType = typeList[idCount]?.ToString() ?? "";

                var writeRow = targetSheet.Dimension.End.Row + 1;

                for (int j = 2; j <= writeCol; j++)
                {
                    var cellTitle = targetSheet.Cells[2, j].Value?.ToString() ?? "";
                    if (cellTitle == "")
                        continue;
                    // 使用 LINQ 查询判断字典中是否包含指定的值
                    bool containsValue = modelSheet.Value.Keys.Any(key =>
                        key.Item1.Equals(itemType) && key.Item2.Equals(cellTitle)
                    );

                    if (containsValue)
                    {
                        var cellModelValue = modelSheet.Value[(itemType, cellTitle)];
                        //分析cellModelValue中的通配符
                        var cellRealValue = AnalyzeWildcard(
                            cellModelValue,
                            exportWildcardData,
                            exportWildcardDyData,
                            baseData,
                            idCount,
                            strDictionary
                        );

                        var cell = targetSheet.Cells[writeRow, j];
                        cell.Value = cellRealValue;
                    }
                }
            }
            targetExcel.Save();
            targetSheet.Dispose();

            NumDesAddIn.App.StatusBar = $"导出：{modelSheetName}";
        }
    }

    private static void TaskSheet(string specialCharsStr) { }

    //分析Cell中通配符构成
    private static string AnalyzeWildcard(
        string cellModelValue,
        Dictionary<string, string> exportWildcardData,
        Dictionary<string, string> exportWildcardDyData,
        Dictionary<string, List<object>> baseData,
        int idCount,
        Dictionary<string, Dictionary<string, List<string>>> strDictionary
    )
    {
        string cellRealValue = cellModelValue;
        string wildcardPattern = "#(.*?)#";
        string wildcardValuePattern = "#";

        MatchCollection matches = Regex.Matches(cellModelValue, wildcardPattern);

        foreach (Match match in matches)
        {
            var wildcard = match.Groups[1].Value;
            if (!exportWildcardData.TryGetValue(wildcard, out var wildcardValue))
            {
                continue;
            }

            var wildcardValueSplit = Regex.Split(wildcardValue, wildcardValuePattern);
            string funName = wildcardValueSplit.ElementAtOrDefault(0) ?? "";
            string funDepends = wildcardValueSplit.ElementAtOrDefault(1) ?? "物品编号";
            string funDy1 = wildcardValueSplit.ElementAtOrDefault(2) ?? "";
            string funDy2 = wildcardValueSplit.ElementAtOrDefault(3) ?? "";
            string funDy3 = wildcardValueSplit.ElementAtOrDefault(4) ?? "";

            string fixWildcardValue = funName switch
            {
                "Left" => Left(exportWildcardDyData, funDepends, funDy1),
                "Right" => Right(exportWildcardDyData, funDepends, funDy1),
                "Set" => Set(exportWildcardDyData, funDepends, funDy1, funDy2),
                "SetDic"
                    => SetDic(
                        exportWildcardDyData,
                        strDictionary,
                        wildcard,
                        funDepends,
                        funDy1,
                        funDy2
                    ),
                "Mer" => Mer(exportWildcardDyData, funDepends, funDy1),
                "MerB" => MerB(exportWildcardDyData, funDepends, funDy1, funDy2, funDy3),
                "Var" => Var(baseData, exportWildcardDyData, wildcard, funDepends, idCount),
                _ => exportWildcardData[wildcard]
            };

            cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
        }

        return cellRealValue;
    }

    private static string Left(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        return exportWildcardDyData[funDepends].Substring(0, int.Parse(funDy1));
    }

    private static string Right(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        return exportWildcardDyData[funDepends]
            .Substring(
                exportWildcardDyData[funDepends].Length - int.Parse(funDy1),
                int.Parse(funDy1)
            );
    }

    private static string Set(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "00" : funDy2;
        return exportWildcardDyData[funDepends]
                .Substring(0, exportWildcardDyData[funDepends].Length - int.Parse(funDy1)) + funDy2;
    }

    private static string SetDic(
        Dictionary<string, string> exportWildcardDyData,
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        string wildcard,
        string funDepends,
        string funDy1,
        string funDy2
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "00" : funDy2;
        string fixWildcardValue = Set(exportWildcardDyData, funDepends, funDy1, funDy2);
        InitializeDictionary(strDictionary, wildcard, fixWildcardValue);
        strDictionary[wildcard][fixWildcardValue].Add(exportWildcardDyData[funDepends]);
        return fixWildcardValue;
    }

    private static string Mer(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        return (long.Parse(exportWildcardDyData[funDepends]) + int.Parse(funDy1)).ToString(); ;
    }
    private static string MerB(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy1) ? "3" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy1) ? "10" : funDy3;
        var baseValue = exportWildcardDyData[funDepends]
            .Substring(exportWildcardDyData[funDepends].Length - 1, 1);
        string result;
        if (int.Parse(baseValue) + int.Parse(funDy1) <= int.Parse(funDy2))
        {
            result = (long.Parse(exportWildcardDyData[funDepends]) + int.Parse(funDy1)).ToString();
        }
        else
        {
            result = (
                long.Parse(exportWildcardDyData[funDepends]) + int.Parse(funDy1) + int.Parse(funDy3)
            ).ToString();
        }
        return result;
    }
    private static string Var(
        Dictionary<string, List<object>> baseData,
        Dictionary<string, string> exportWildcardDyData,
        string wildcard,
        string funDepends,
        int idCount
    )
    {
        string fixWildcardValue = baseData[funDepends][idCount].ToString();
        exportWildcardDyData[wildcard] = fixWildcardValue;
        return fixWildcardValue;
    }

    //自定义字典初始化
    private static void InitializeDictionary(
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        string key,
        string subKey
    )
    {
        if (!strDictionary.ContainsKey(key))
        {
            strDictionary[key] = new Dictionary<string, List<string>>();
        }
        if (!strDictionary[key].ContainsKey(subKey))
        {
            strDictionary[key][subKey] = new List<string>();
        }
    }
}
