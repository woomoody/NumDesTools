using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using Match = System.Text.RegularExpressions.Match;

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
        //需要查找通配符字段所在列
        var exportWildcardRange = ws.Range["A1:AZ1"].Value2;
        var exportWildcardCol = PubMetToExcel.FindValueIn2DArray(exportWildcardRange, "通配符").Item2;

        //基本信息
        var exportBaseSheetData = new Dictionary<string, Dictionary<string, Tuple<int, int>>>();
        var exportBaseData = new Dictionary<string, Tuple<int, int>>();
        if (exportBaseData == null)
            throw new ArgumentNullException(nameof(exportBaseData));

        object[,] baseRangeValue = ws.Range[
            ws.Cells[selectRow, selectCol + 2],
            ws.Cells[selectRow + 2, exportWildcardCol - 2]
        ].Value2;

        for (int col = 1; col <= baseRangeValue.GetLength(1); col++)
        {
            var keyName = baseRangeValue[1, col]?.ToString() ?? "";
            if (keyName != "")
            {
                if (baseRangeValue[2, col] == null)
                {
                    continue;
                }
                var keyCol = Convert.ToInt32(baseRangeValue[2, col]);
                var keyRowMax = Convert.ToInt32(baseRangeValue[3, col]);
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
            ws.Cells[selectRow, exportWildcardCol],
            ws.Cells[selectRow + wildcardCount, exportWildcardCol + 1]
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

        string id;
        string idType;

        if (baseSheetName.Contains("【基础】"))
        {
  
                //走【基础】表逻辑
                id = "数据编号";
                idType = "类型";
                BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);
        }
        else if (baseSheetName.Contains("【任务】"))
        {
            //走【基础】表逻辑
            id = "任务编号";
            idType = "类型";
            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);
        }
        else if (baseSheetName.Contains("【坐标】"))
        {
            //走【基础】表逻辑
            id = "编号";
            idType = "类型";
            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);
        }

        else if (baseSheetName.Contains("【通用】"))
        {
            //走【基础】表逻辑
            id = "数据编号";
            idType = "类型";
            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);
        }
        sw.Stop();
        var ts2 = sw.Elapsed;
        NumDesAddIn.App.StatusBar = "导出完成，用时：" + ts2;
    }

    //个别导出LTE数据配置
    public static void ExportLteDataConfigSelf(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sw = new Stopwatch();
        sw.Start();

        //获取【导出】表信息
        var ws = Wk.ActiveSheet;
        var selectRange = NumDesAddIn.App.Selection;
        var baseSheetName = selectRange.Value2.ToString();
        var selectRow = selectRange.Row;
        var selectCol = selectRange.Column;
        //需要查找通配符字段所在列
        var exportWildcardRange = ws.Range["A1:AZ1"].Value2;
        var exportWildcardCol = PubMetToExcel.FindValueIn2DArray(exportWildcardRange, "通配符").Item2;

        //基本信息
        var exportBaseSheetData = new Dictionary<string, Dictionary<string, Tuple<int, int>>>();
        var exportBaseData = new Dictionary<string, Tuple<int, int>>();
        if (exportBaseData == null)
            throw new ArgumentNullException(nameof(exportBaseData));

        object[,] baseRangeValue = ws.Range[
            ws.Cells[selectRow, selectCol + 2],
            ws.Cells[selectRow + 2, exportWildcardCol - 2]
        ].Value2;

        for (int col = 1; col <= baseRangeValue.GetLength(1); col++)
        {
            var keyName = baseRangeValue[1, col]?.ToString() ?? "";
            if (keyName != "")
            {
                if (baseRangeValue[2, col] == null)
                {
                    continue;
                }
                var keyCol = Convert.ToInt32(baseRangeValue[2, col]);
                var keyRowMax = Convert.ToInt32(baseRangeValue[3, col]);
                exportBaseData[keyName] = new Tuple<int, int>(keyCol, keyRowMax);
            }
        }
        exportBaseSheetData[baseSheetName] = exportBaseData;

        //通配符信息
        var exportWildcardData = new Dictionary<string, string>();
        if (exportWildcardData == null)
            throw new ArgumentNullException(nameof(exportWildcardData));

        int wildcardCount;
        if (ws.Cells[selectRow + 1, selectCol].Value2 == null)
        {
            MessageBox.Show("需要选中表格名的单元格，通配符数量为空，请检查数据");
            return;
        }
        else
        {
            wildcardCount = (int)ws.Cells[selectRow + 1, selectCol].Value2;
        }

        var wildcardRangeValue = ws.Range[
            ws.Cells[selectRow, exportWildcardCol],
            ws.Cells[selectRow + wildcardCount, exportWildcardCol + 1]
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

        string id;
        string idType;

        if (baseSheetName.Contains("【基础】"))
        {

            //走【基础】表逻辑
            id = "数据编号";
            idType = "类型";

            var keysToFilter = GetCellValuesFromUserInput("【基础】");
            if (keysToFilter != null)
            {
                baseData = FilterBySpecifiedKeyAndSyncPositions(baseData, id, keysToFilter);
            }
            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);

        }
        else if (baseSheetName.Contains("【任务】"))
        {
            //走【基础】表逻辑
            id = "任务编号";
            idType = "类型";
            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);
        }
        else if (baseSheetName.Contains("【坐标】"))
        {
            //走【基础】表逻辑
            id = "编号";
            idType = "类型";
            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);
        }

        else if (baseSheetName.Contains("【通用】"))
        {
            //走【基础】表逻辑
            id = "数据编号";
            idType = "类型";
            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType);
        }
        sw.Stop();
        var ts2 = sw.Elapsed;
        NumDesAddIn.App.StatusBar = "导出完成，用时：" + ts2;
    }

    private static Dictionary<string, List<object>> FilterBySpecifiedKeyAndSyncPositions(
        Dictionary<string, List<object>> baseData,
        string targetKey,
        List<string> cellValues)
    {
        // 如果 baseData 不包含目标 Key，直接返回空字典
        if (!baseData.ContainsKey(targetKey))
        {
            return new Dictionary<string, List<object>>();
        }

        // 获取目标 Key 的 List
        List<object> targetList = baseData[targetKey];

        // 转换为 HashSet 提高性能
        var valueSet = new HashSet<string>(cellValues);

        // 找出目标 List 中符合条件的元素的索引位置
        List<int> matchedIndices = targetList
            .Select((item, index) => new { item, index })
            .Where(x => valueSet.Contains(x.item.ToString()))
            .Select(x => x.index)
            .ToList();

        // 如果没有匹配项，返回空字典
        if (matchedIndices.Count == 0)
        {
            return new Dictionary<string, List<object>>();
        }

        // 构建筛选后的新 baseData
        var filteredData = new Dictionary<string, List<object>>();
        foreach (var kv in baseData)
        {
            // 对每个 Key 的 List，只保留 matchedIndices 对应的元素
            List<object> filteredList = matchedIndices
                .Select(i => kv.Value[i])
                .ToList();

            filteredData.Add(kv.Key, filteredList);
        }

        return filteredData;
    }

    //获取用户输入的单元格值
    private static  List<string> GetCellValuesFromUserInput(string sheetName)
    {
        Range selectedRange = NumDesAddIn.App.InputBox(
            $"请用鼠标选择{sheetName}单元格（Ctr，可多选）",
            "选择单元格",
            Type: 8
        ) as Range;

        if (selectedRange == null)
        {
            MessageBox.Show("未选择任何单元格！");
            return null;
        }

        // 遍历所选单元格，获取值
        List<string> cellValues = new List<string>();
        foreach (Range cell in selectedRange)
        {
            try
            {
                string value = cell.Value?.ToString();
                cellValues.Add(value ?? "");
            }
            catch (Exception ex)
            {
                cellValues.Add($"错误: 无法读取 {cell.Address} - {ex.Message}");
            }
        }

        return cellValues;
    }
    private static void BaseSheet(
        Dictionary<string, List<object>> baseData,
        Dictionary<string, string> exportWildcardData,
        Dictionary<string, Dictionary<(object, object), string>> modelValueAll,
        string id,
        string idType
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var strDictionary = new Dictionary<string, Dictionary<string, List<string>>>();

        var idList = baseData[id];
        var typeList = baseData[idType];

        //替换通配符生成数据

        foreach (var modelSheet in modelValueAll)
        {
            string modelSheetName = modelSheet.Key;

            PubMetToExcel.SetExcelObjectEpPlus(
                WkPath,
                modelSheetName,
                out ExcelWorksheet targetSheet,
                out ExcelPackage targetExcel
            );

            if (targetSheet == null)
            {
                LogDisplay.RecordLine(
                    "[{0}] , {1}【#LTE数据模版】中创建的文件名不存在",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    modelSheetName
                );
            }

            if (targetSheet != null)
            {
                NumDesAddIn.App.StatusBar = $"导出：{modelSheetName}";

                var writeCol = targetSheet.Dimension.End.Column;

                var exportWildcardDyData = new Dictionary<string, string>(exportWildcardData);

                bool dataWritten = false; // 标志是否有实际写入
                var dataRepeatWritten = new HashSet<string>();

                for (int idCount = 0; idCount < idList.Count; idCount++)
                {
                    string itemId = idList[idCount]?.ToString() ?? "";
                    if (itemId == "")
                        continue;
                    string itemType = typeList[idCount]?.ToString() ?? "";

                    var writeRow = targetSheet.Dimension.End.Row + 1;
                    //更新动态值
                    foreach (var wildcardDy in exportWildcardData)
                    {
                        GetDyWildcardValue(
                            baseData,
                            exportWildcardDyData,
                            wildcardDy.Key,
                            wildcardDy.Value,
                            idCount
                        );
                    }

           

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
                                strDictionary
                            );

                            //空ID判断
                            if (j == 2 && cellRealValue == string.Empty)
                            {
                                break;
                            }
                            //重复ID判断
                            if (j == 2 && dataRepeatWritten.Contains(cellRealValue))
                            {
                                break;
                            }

                            if (j == 2)
                            {
                                //字典型数据判断，需要数据计算完毕后单独写入
                                dataRepeatWritten.Add(cellRealValue);
                            }
                            //实际写入
                            var cell = targetSheet.Cells[writeRow, j];
                            cell.Value = cellRealValue;
                            dataWritten = true;

                        }
                    }
                }
                if (dataWritten) // 只有在写入数据时才保存
                {
                    targetExcel.Save();
                }
            }

            if (targetSheet != null) targetSheet.Dispose();
        }
        //输出字典数据
        if (strDictionary.Count > 0)
        {
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Path.Combine(documentsPath, "strDic.csv");
            SaveDictionaryToFile(strDictionary, filePath);
        }

    }

    //分析Cell中通配符构成
    private static string AnalyzeWildcard(
        string cellModelValue,
        Dictionary<string, string> exportWildcardData,
        Dictionary<string, string> exportWildcardDyData,
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
                //根据动态或静态值计算值
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
                "Ads" => Ads(exportWildcardDyData, funDepends, funDy1),
                "Arr" => Arr(exportWildcardDyData, funDepends, funDy1, funDy2),
                "Get" => Get(exportWildcardDyData, funDepends, funDy1, funDy2),
                "GetDic" => GetDic(strDictionary, exportWildcardDyData , funDepends , funDy1 , funDy2, funDy3),
                "GetDicKey" => GetDicKey(funDepends),
                //获取动态值
                "Var" => exportWildcardDyData[wildcard],
                //获取静态值
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
        var maxCount = Math.Min(exportWildcardDyData[funDepends].Length, int.Parse(funDy1));
        return exportWildcardDyData[funDepends].Substring(0, maxCount);
    }

    private static string Right(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        var maxCount = Math.Min(exportWildcardDyData[funDepends].Length, int.Parse(funDy1));
        return exportWildcardDyData[funDepends]
            .Substring(exportWildcardDyData[funDepends].Length - maxCount, int.Parse(funDy1));
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
        return (long.Parse(exportWildcardDyData[funDepends]) + int.Parse(funDy1)).ToString();
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
        funDy2 = string.IsNullOrEmpty(funDy2) ? "3" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "10" : funDy3;
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

    private static string Ads(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "链类最大值" : funDy1;
        string rootNum =
            exportWildcardDyData[funDepends]
                .Substring(0, exportWildcardDyData[funDepends].Length - 2) + "00";
        int baseValue = int.Parse(
            exportWildcardDyData[funDepends]
                .Substring(exportWildcardDyData[funDepends].Length - 1, 1)
        );
        int baseMax = 0;
        try
        {
            baseMax = int.Parse(exportWildcardDyData[funDy1]);
        }
        catch (Exception e)
        {
            MessageBox.Show($"{rootNum}##{funDy1}可能为空{e.Message}");
        }
        if (baseMax == 0)
        {
            MessageBox.Show($"{rootNum}物品应该不属于链");
        }
        var loopNum = LoopNumber(baseValue, baseMax);
        string result = "";
        foreach (var num in loopNum)
        {
            var digNum = (long.Parse(rootNum) + num).ToString();
            result += digNum + ",";
        }

        result = result.Substring(0, result.Length - 1);
        result = $"{result}";

        return result;
    }

    private static string Arr(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "消耗量组" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "" : funDy2;

        var funDy1Value = exportWildcardDyData[funDy1];
        var funDependsValue = exportWildcardDyData[funDepends];

        var funDy1ValueSplit = Regex.Split(funDy1Value, ",");

        var funDependsValueSplit = Regex.Split(funDependsValue, ",");

        string result = "";
        if (funDy1ValueSplit.Length == funDependsValueSplit.Length)
        {
            for (int i = 0; i < funDy1ValueSplit.Length; i++)
            {
                string temp;
                if (funDy2 != "")
                {
                    var funDy2Value = exportWildcardDyData[funDy2];
                    if (long.TryParse(funDy2Value, out long funDy2ValueLong))
                    {
                        temp =
                            $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]},{funDy2ValueLong + i}]";
                    }
                    else
                    {
                        temp = $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]},{funDy2Value}]";
                    }
                }
                else
                {
                    temp = $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]}]";
                }
                result += temp + ",";
            }
            result = result.Substring(0, result.Length - 1);
        }
        return result;
    }

    private static string Get(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "," : funDy2;
        var dependsValue = exportWildcardDyData[funDepends];
        var dependsValueSplit = Regex.Split(dependsValue, funDy2);
        var result = dependsValueSplit[int.Parse(funDy1) - 1];
        return result;
    }

    private static string GetDic(
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "物品编号" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "2" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "00" : funDy1;
        var baseDicKey = exportWildcardDyData[funDy1]
            .Substring(0, exportWildcardDyData[funDy1].Length - int.Parse(funDy2)) + funDy3;
        var dependsDicValue = strDictionary[funDepends];
        var dependsValueList = dependsDicValue[baseDicKey];
        // 去重并按照数字从小到大排序
        List<string> distinctSortedNumbers = dependsValueList
            .Distinct()                      // 去重
            .OrderBy(n => long.Parse(n))       // 按照数字从小到大排序
            .ToList();                        // 转换回 List
        return string.Join(",", distinctSortedNumbers);
    }
    private static string GetDicKey(
        string funDepends
    )
    {
        //读取本地存储数据
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string filePath = Path.Combine(documentsPath, "strDic.csv");
        var fileDicData = LoadDictionaryFromFile(filePath);
        var dependsDicValue = fileDicData[funDepends];
        return string.Join(",", dependsDicValue.Keys);
    }
    //获取动态值
    private static void GetDyWildcardValue(
        Dictionary<string, List<object>> baseData,
        Dictionary<string, string> exportWildcardDyData,
        string wildcard,
        string funDepends,
        int idCount
    )
    {
        if (funDepends.Contains("Var"))
        {
            var wildcardValueSplit = Regex.Split(funDepends, "#");
            string fixWildcardValue = baseData[wildcardValueSplit[1]][idCount]?.ToString() ?? "";
            //ID组关键词替换
            if (wildcardValueSplit.Length == 3)
            {
                fixWildcardValue = fixWildcardValue.Replace("#", wildcardValueSplit[2]);
            }
            exportWildcardDyData[wildcard] = fixWildcardValue;
        }
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
            strDictionary[key][subKey] = [];
        }
    }

    //循环数字
    private static List<int> LoopNumber(int start, int max)
    {
        List<int> sequence = [];

        for (int i = 1; i <= max; i++)
        {
            var modValue = ((start - 1) % max) + 1;
            start++;
            sequence.Add(modValue);
        }

        return sequence;
    }

    //strDic输出到文件
    private static void SaveDictionaryToFile(
        Dictionary<string, Dictionary<string, List<string>>> dictionary,
        string filePath
    )
    {
        using StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8);
        foreach (var outerPair in dictionary)
        {
            foreach (var innerPair in outerPair.Value)
            {
                var line = $"{outerPair.Key},{innerPair.Key},{string.Join(",", innerPair.Value)}";
                writer.WriteLine(line);
            }
        }
    }
    //文件输出到strDic
    private static Dictionary<string, Dictionary<string, List<string>>> LoadDictionaryFromFile(string filePath)
    {
        var dictionary = new Dictionary<string, Dictionary<string, List<string>>>();

        using StreamReader reader = new StreamReader(filePath, Encoding.UTF8);
        string line;
        while ((line = reader.ReadLine()) != null)
        {
            // 拆分每一行，假设格式为 outerKey,innerKey,value1,value2,...
            var parts = line.Split(',');

            if (parts.Length < 2)
            {
                // 如果行的格式不正确，跳过该行
                continue;
            }

            string outerKey = parts[0];
            string innerKey = parts[1];
            List<string> values = new List<string>(parts[2..]); // 从第三个元素开始是 values

            // 如果外层字典中没有 outerKey，先创建一个新的字典
            if (!dictionary.ContainsKey(outerKey))
            {
                dictionary[outerKey] = new Dictionary<string, List<string>>();
            }

            // 将 innerKey 和 values 添加到内层字典中
            dictionary[outerKey][innerKey] = values;
        }

        return dictionary;
    }
}
