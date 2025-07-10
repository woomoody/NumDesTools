using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Input;
using OfficeOpenXml;
using Match = System.Text.RegularExpressions.Match;

namespace NumDesTools;

public class LteData
{
    private static readonly Workbook Wk = NumDesAddIn.App.ActiveWorkbook;

    private static readonly string WkPath = Wk.Path;

    private static readonly Regex WildcardRegex = new("#(.*?)#", RegexOptions.Compiled);

    private static readonly Dictionary<string, (string id, string idType)> SheetTypeMap =
        new(StringComparer.Ordinal)
        {
            ["LTE【基础】"] = ("数据编号", "类型"),
            ["LTE【任务】"] = ("任务编号", "类型"),
            ["LTE【寻找】"] = ("寻找编号", "类型"),
            ["LTE【通用】"] = ("数据编号", "类型")
        };

    private const int BaseDataTagCol = 0;
    private const int BaseDataStartCol = 1;
    private const int BaseDataEndCol = 33;
    private const int FindDataTagCol = 0;
    private const int FindDataStartCol = 1;
    private const int FindDataEndCol = 9;
    private const int TaskDataTagCol = 13;
    private const int TaskDataStartCol = 14;
    private const int TaskDataEndCol = 23;

    private const string ActivityIdIndex = "B1";
    private const string ActivityDataMinIndex = "C1";
    private const string ActivityDataMaxIndex = "D1";
    private const string ActivityNameMinIndex = "E1";

    private static readonly Worksheet PubWildSheet = Wk.Worksheets["LTE【设计】"];
    private static readonly Dictionary<string, string> OutputWildcardPubDic =
        new()
        {
            ["活动编号"] = (PubWildSheet.Range[ActivityIdIndex].Value2 / 10000)?.ToString(),
            ["活动备注"] = PubWildSheet.Range[ActivityNameMinIndex].Value2?.ToString(),
        };

    #region LTE数据配置导出
    //导出LTE数据配置
    public static void ExportLteDataConfigFirst(CommandBarButton ctrl, ref bool cancelDefault)
    {
        ExportLteDataConfig(true);
    }

    public static void ExportLteDataConfigUpdate(CommandBarButton ctrl, ref bool cancelDefault)
    {
        ExportLteDataConfig(false);
    }

    public static void ExportLteDataConfig(bool isFirst)
    {
        //Epplus获取[LTE配置【导出】]表的ListObject
        var sheet = Wk.ActiveSheet;
        var sheetName = sheet.Name;
        var outputData = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "LTE配置【导出】"
        );

        var outputWildcardArray = outputData[$"{sheetName}_通配符"];
        var outputWildcardDic = PubMetToExcel.TwoDArrayToDicFirstKeyStr(outputWildcardArray);
        var outputWildDic = outputWildcardDic
            .Concat(OutputWildcardPubDic)
            .GroupBy(k => k.Key)
            .ToDictionary(g => g.Key, g => g.Last().Value);

        //Epplus获取[#LTE数据模版]表的ListObject
        if (GetModelValue(out var modelValueAll, "#LTE数据模版"))
        {
            return;
        }

        if (SheetTypeMap.ContainsKey(sheetName))
        {
            var kv = SheetTypeMap[sheetName];
            string id = kv.Item1;
            string idType = kv.Item2;
            //获取【当前表】ListObject
            var sheetListObject = sheet.ListObjects[sheetName];
            if (sheetListObject == null)
            {
                MessageBox.Show($"{sheetName}不存在{sheetName}的名称，请检查");
                return;
            }

            Range sheetListObjectRange = sheetListObject.Range;

            // 获取基础数据
            object[,] sheetBaseArray = sheetListObjectRange.Value2;

            // 获取更新标识数据
            int isFirstTagStartRow = sheetListObjectRange.Row;
            int isFirstTagStartCol = sheetListObjectRange.Column;
            int isFirstTagEndRow = sheetListObjectRange.Rows.Count - 1 + isFirstTagStartRow;
            
            Range firstTagRange = sheet.Range[
                sheet.Cells[isFirstTagStartRow, isFirstTagStartCol -1],
                sheet.Cells[isFirstTagEndRow, isFirstTagStartCol -1]
            ];

            object[,] firstTagArray = firstTagRange.Value2;

            // 合并数据
            var mergeBaseArray = PubMetToExcel.Merge2DArrays1(sheetBaseArray , firstTagArray);

            var sheetBaseDic = PubMetToExcel.TwoDArrayToDictionaryFirstRowKey(mergeBaseArray);
            //执行导出逻辑
            BaseSheet(sheetBaseDic, outputWildDic, modelValueAll, id, idType, isFirst);
        }
        else
        {
            MessageBox.Show($"{sheetName}有误，请对比#【A-LTE】配置模版");
        }
    }

    private static bool GetModelValue(
        out Dictionary<string, Dictionary<(object, object), string>> modelValueAll,
        string sheetName
    )
    {
        PubMetToExcel.SetExcelObjectEpPlusNormal(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            sheetName,
            out ExcelWorksheet modelSheet,
            out ExcelPackage modelExcel
        );

        modelValueAll = new Dictionary<string, Dictionary<(object, object), string>>();
        foreach (var table in modelSheet.Tables)
        {
            if (table != null)
            {
                var modeName = table.Name;

                object[,] tableData =
                    modelSheet
                        .Cells[
                            table.Address.Start.Row,
                            table.Address.Start.Column,
                            table.Address.End.Row,
                            table.Address.End.Column
                        ]
                        .Value as object[,];

                if (tableData is not null)
                {
                    int rowCount = tableData.GetLength(0);
                    int colCount = tableData.GetLength(1);

                    // 将二维数组的数据存储到字典中
                    var modelValue = PubMetToExcel.Array2DToDic2D0(rowCount, colCount, tableData);
                    if (modelValue == null)
                    {
                        return true;
                    }
                    modelValueAll[modeName] = modelValue;
                }
            }
            else
            {
                Debug.Print("表格不存在");
            }
        }
        modelExcel?.Dispose();
        return false;
    }

    //个别导出LTE数据配置
    //public static void ExportLteDataConfigSelf(CommandBarButton ctrl, ref bool cancelDefault)
    //{
    //    NumDesAddIn.App.StatusBar = false;
    //    var sw = new Stopwatch();
    //    sw.Start();

    //    //读取【基础/任务……】表数据
    //    var sheetInfo = ReadExportSheetInfo();
    //    string baseSheetName = sheetInfo.baseSheetName;
    //    var baseSheet = Wk.Worksheets[baseSheetName];
    //    var baseData = new Dictionary<string, List<object>>();
    //    var baseSheetData = sheetInfo.exportBaseData;
    //    var exportWildcardData = sheetInfo.exportWildcardData;

    //    foreach (var baseElement in baseSheetData)
    //    {
    //        var range = baseSheet
    //            .Range[
    //                baseSheet.Cells[2, baseElement.Value.Item1],
    //                baseSheet.Cells[baseElement.Value.Item2, baseElement.Value.Item1]
    //            ]
    //            .Value2;

    //        var dataList = PubMetToExcel.List2DToListRowOrCol(
    //            PubMetToExcel.RangeDataToList(range),
    //            true
    //        );

    //        baseData[baseElement.Key] = dataList;
    //    }

    //    //Epplus获取[#LTE数据模版]表的ListObject
    //    if (GetModelValue(out var modelValueAll, "#LTE数据模版"))
    //    {
    //        return;
    //    }

    //    string id;
    //    string idType;

    //    foreach (var kv in SheetTypeMap)
    //    {
    //        if (baseSheetName.Contains(kv.Key))
    //        {
    //            id = kv.Value.id;
    //            idType = kv.Value.idType;

    //            //自选新增数据，否则全量数据
    //            var keysToFilter = GetCellValuesFromUserInput(kv.Key);
    //            if (keysToFilter != null)
    //            {
    //                baseData = FilterBySpecifiedKeyAndSyncPositions(baseData, id, keysToFilter);
    //            }

    //            BaseSheet(baseData, exportWildcardData, modelValueAll, id, idType, false);
    //            break;
    //        }
    //    }

    //    sw.Stop();
    //    var ts2 = sw.ElapsedMilliseconds;
    //    NumDesAddIn.App.StatusBar = "导出完成，用时：" + ts2;
    //}

    //private static (
    //    string baseSheetName,
    //    int selectRow,
    //    int selectCol,
    //    Dictionary<string, Tuple<int, int>> exportBaseData,
    //    Dictionary<string, string> exportWildcardData
    //) ReadExportSheetInfo()
    //{
    //    Worksheet ws = Wk.ActiveSheet;
    //    var selectRange = NumDesAddIn.App.Selection;
    //    string baseSheetName = selectRange.Value2.ToString();
    //    int selectRow = selectRange.Row;
    //    int selectCol = selectRange.Column;

    //    // 查找通配符列
    //    object[,] exportWildcardRange = ws.Range["A1:AZ1"].Value2;
    //    int exportWildcardCol = PubMetToExcel.FindValueIn2DArray(exportWildcardRange, "通配符").Item2;

    //    // 读取基础数据
    //    var exportBaseData = new Dictionary<string, Tuple<int, int>>();
    //    object[,] baseRangeValue = ws.Range[
    //        ws.Cells[selectRow, selectCol + 2],
    //        ws.Cells[selectRow + 2, exportWildcardCol - 2]
    //    ].Value2;

    //    for (int col = 1; col <= baseRangeValue.GetLength(1); col++)
    //    {
    //        string keyName = baseRangeValue[1, col]?.ToString() ?? "";
    //        if (!string.IsNullOrEmpty(keyName) && baseRangeValue[2, col] != null)
    //        {
    //            int keyCol = Convert.ToInt32(baseRangeValue[2, col]);
    //            int keyRowMax = Convert.ToInt32(baseRangeValue[3, col]);
    //            exportBaseData[keyName] = Tuple.Create(keyCol, keyRowMax);
    //        }
    //    }

    //    // 读取通配符数据
    //    var exportWildcardData = new Dictionary<string, string>();
    //    int wildcardCount = (int)ws.Cells[selectRow + 1, selectCol].Value2;
    //    object[,] wildcardRangeValue = ws.Range[
    //        ws.Cells[selectRow, exportWildcardCol],
    //        ws.Cells[selectRow + wildcardCount, exportWildcardCol + 1]
    //    ].Value2;

    //    for (int row = 1; row <= wildcardCount; row++)
    //    {
    //        string wildcardName = wildcardRangeValue[row, 1]?.ToString() ?? "";
    //        if (!string.IsNullOrEmpty(wildcardName))
    //        {
    //            exportWildcardData[wildcardName] = wildcardRangeValue[row, 2].ToString();
    //        }
    //    }

    //    return (baseSheetName, selectRow, selectCol, exportBaseData, exportWildcardData);
    //}

    //private static Dictionary<string, List<object>> FilterBySpecifiedKeyAndSyncPositions(
    //    Dictionary<string, List<object>> baseData,
    //    string targetKey,
    //    List<string> cellValues
    //)
    //{
    //    // 如果 baseData 不包含目标 Key，直接返回空字典
    //    if (!baseData.ContainsKey(targetKey))
    //    {
    //        return new Dictionary<string, List<object>>();
    //    }

    //    // 获取目标 Key 的 List
    //    List<object> targetList = baseData[targetKey];

    //    // 转换为 HashSet 提高性能
    //    var valueSet = new HashSet<string>(cellValues);

    //    // 找出目标 List 中符合条件的元素的索引位置
    //    List<int> matchedIndices = targetList
    //        .Select((item, index) => new { item, index })
    //        .Where(x => valueSet.Contains(x.item.ToString()))
    //        .Select(x => x.index)
    //        .ToList();

    //    // 如果没有匹配项，返回空字典
    //    if (matchedIndices.Count == 0)
    //    {
    //        return new Dictionary<string, List<object>>();
    //    }

    //    // 构建筛选后的新 baseData
    //    var filteredData = new Dictionary<string, List<object>>();
    //    foreach (var kv in baseData)
    //    {
    //        // 对每个 Key 的 List，只保留 matchedIndices 对应的元素
    //        List<object> filteredList = matchedIndices.Select(i => kv.Value[i]).ToList();

    //        filteredData.Add(kv.Key, filteredList);
    //    }

    //    return filteredData;
    //}

    ////获取用户输入的单元格值
    //private static List<string> GetCellValuesFromUserInput(string sheetName)
    //{
    //    Range selectedRange =
    //        NumDesAddIn.App.InputBox($"请用鼠标选择{sheetName}单元格（Ctr，可多选）", "选择单元格", Type: 8) as Range;

    //    if (selectedRange == null)
    //    {
    //        MessageBox.Show("未选择任何单元格！");
    //        return null;
    //    }

    //    // 遍历所选单元格，获取值
    //    List<string> cellValues = new List<string>();
    //    foreach (Range cell in selectedRange)
    //    {
    //        try
    //        {
    //            string value = cell.Value?.ToString();
    //            cellValues.Add(value ?? "");
    //        }
    //        catch (Exception ex)
    //        {
    //            cellValues.Add($"错误: 无法读取 {cell.Address} - {ex.Message}");
    //        }
    //    }

    //    return cellValues;
    //}

    private static void BaseSheet(
        Dictionary<string, List<string>> baseData,
        Dictionary<string, string> exportWildcardData,
        Dictionary<string, Dictionary<(object, object), string>> modelValueAll,
        string id,
        string idType,
        bool isFirst = true
    )
    {
        var strDictionary = new Dictionary<string, Dictionary<string, List<string>>>();

        var idList = baseData[id];
        var typeList = baseData[idType];

        //过滤id和type，只针对有增删改的数据进行导出
        List<string> dataStatusList = null;
        List<string> dataStatusListNew = null;
        if (!isFirst)
        {
            if (baseData.ContainsKey("数据状态"))
            {
                dataStatusList = baseData["数据状态"];
            }
            if (dataStatusList != null)
            {
                var combined = idList
                    .Zip(typeList, (dataId, type) => new { id = dataId, type })
                    .Zip(
                        dataStatusList,
                        (first, status) =>
                            new
                            {
                                first.id,
                                first.type,
                                status
                            }
                    )
                    .Where(x => x.status?.ToString() is "+" or "-" or "*")
                    .ToList();

                idList = combined.Select(x => x.id).ToList();
                typeList = combined.Select(x => x.type).ToList();
                dataStatusListNew = combined.Select(x => x.status).ToList();
            }
        }

        //替换通配符生成数据
        foreach (var modelSheet in modelValueAll)
        {
            // 同名表格需要导出多套数据时，需要过滤正确的文件名称
            string modelSheetName = Regex.Replace(modelSheet.Key, @"No\d+", "");

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
                    string itemId = idList[idCount] ?? "";
                    if (itemId == "")
                        continue;
                    string itemType = typeList[idCount] ?? "";

                    var writeRow = targetSheet.Dimension.End.Row + 1;

                    //非第一次执行更新逻辑，否则首次逻辑
                    if (dataStatusListNew != null)
                    {
                        //查找ID是否存在
                        var rowIndex = PubMetToExcel.FindSourceRow(targetSheet, 2, itemId);
                        //如果存在且标识为删除，则删除，不进行写入，标识为修改则进行写入
                        if (rowIndex != -1)
                        {
                            if (dataStatusListNew[idCount] == "-")
                            {
                                targetSheet.DeleteRow(rowIndex);
                                dataWritten = true;
                                continue;
                            }

                            if (dataStatusListNew[idCount] == "*")
                            {
                                writeRow = rowIndex;
                            }
                            else
                            {
                                //带#的已经写入过，不需要再写入
                                continue;
                            }
                        }
                        //如果不存在，则需要寻找本表相似ID最大行，依次写入
                        else
                        {
                            if (dataStatusListNew[idCount] == "-")
                            {
                                //跳过标记为删除，但目标表也不存在的数据
                                continue;
                            }

                            var activeId = itemId.Substring(0, 6);
                            var regexSearch = new Regex($"^{activeId}\\d{{4}}$");
                            var baseWriteRow = PubMetToExcel.FindSourceRowBlur(
                                targetSheet,
                                2,
                                regexSearch
                            );
                            if (baseWriteRow != -1)
                            {
                                if (writeRow != baseWriteRow + 1)
                                {
                                    //需要插入行
                                    targetSheet.InsertRow(baseWriteRow, 1);
                                    writeRow = baseWriteRow + 1;
                                }
                            }
                        }
                    }

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

                    //整理写入数据
                    var writeData = new Dictionary<(int row, int col), string>();
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
                            // 分析cellModelValue中的通配符
                            var cellRealValue = AnalyzeWildcard(
                                cellModelValue,
                                exportWildcardData,
                                exportWildcardDyData,
                                strDictionary,
                                baseData,
                                id,
                                itemId
                            );

                            // 空ID判断
                            if (j == 2 && cellRealValue == string.Empty)
                            {
                                break;
                            }
                            // 重复ID判断
                            if (j == 2 && dataRepeatWritten.Contains(cellRealValue))
                            {
                                break;
                            }

                            if (j == 2)
                            {
                                // 字典型数据判断，需要数据计算完毕后单独写入
                                dataRepeatWritten.Add(cellRealValue);
                            }
                            // 记录数据
                            var cell = targetSheet.Cells[writeRow, j];
                            if (isFirst)
                            {
                                writeData[(writeRow, j)] = cellRealValue;
                                dataWritten = true;
                            }
                            else
                            {
                                if (cell.Value?.ToString() != cellRealValue)
                                {
                                    writeData[(writeRow, j)] = cellRealValue;
                                    dataWritten = true;
                                }
                            }
                        }
                    }
                    // 实际写入
                    foreach (var cell in writeData)
                    {
                        targetSheet.Cells[cell.Key.row, cell.Key.col].Value = cell.Value;
                    }
                }
                if (dataWritten) // 只有在写入数据时才保存
                {
                    targetExcel.Save();
                }
            }

            targetExcel?.Dispose();
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
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        Dictionary<string, List<string>> baseData,
        string id,
        string itemId
    )
    {
        string cellRealValue = cellModelValue;
        string wildcardValuePattern = "#";
        List<string> idList = baseData[id];

        MatchCollection matches = WildcardRegex.Matches(cellModelValue);

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
            string funDy4 = wildcardValueSplit.ElementAtOrDefault(5) ?? "";
            string funDy5 = wildcardValueSplit.ElementAtOrDefault(6) ?? "";

            try
            {
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
                            funDy2,
                            funDy3
                        ),
                    "Mer" => Mer(exportWildcardDyData, funDepends, itemId, funDy1),
                    "MerB"
                        => MerB(exportWildcardDyData, funDepends, itemId, funDy1, funDy2, funDy3),
                    "MerTry"
                        => MerTry(exportWildcardDyData, funDepends, funDy1, funDy2, funDy3, idList),
                    "Ads" => Ads(exportWildcardDyData, funDepends, funDy1),
                    "Arr" => Arr(exportWildcardDyData, funDepends, funDy1, funDy2),
                    "Get" => Get(exportWildcardDyData, funDepends, funDy1, funDy2),
                    "GetDic"
                        => GetDic(
                            strDictionary,
                            exportWildcardDyData,
                            funDepends,
                            funDy1,
                            funDy2,
                            funDy3
                        ),
                    "GetDicKey" => GetDicKey(funDepends),
                    "SplitArr" => SplitArr(exportWildcardDyData, funDepends, funDy1, funDy2),
                    "CollectRow"
                        => CollectRow(
                            exportWildcardDyData,
                            funDepends,
                            funDy1,
                            funDy2,
                            funDy3,
                            funDy4,
                            funDy5,
                            baseData,
                            id
                        ),
                    //获取动态值
                    "Var" => exportWildcardDyData[wildcard],

                    //获取静态值
                    _ => exportWildcardData[wildcard]
                };

                cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
            }
            catch (FormatException)
            {
                Debug.Print($"通配符解析错误: {wildcard} | 值: {exportWildcardDyData[wildcard]}");
                LogDisplay.RecordLine(
                    "[{1}][{0}][{2}]",
                    $"通配符解析错误: {wildcard} | 值: {exportWildcardDyData[wildcard]}",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    wildcard
                );
                return string.Empty;
            }
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
        string funDy2,
        string funDy3
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "00" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "链类最大值" : funDy3;

        string fixWildcardValue = Set(exportWildcardDyData, funDepends, funDy1, funDy2);
        InitializeDictionary(strDictionary, wildcard, fixWildcardValue);

        string maxLink = exportWildcardDyData[funDy3];
        if(maxLink !="")
        {
            var linkList = new List<string>();
            for (int i = 0; i < int.Parse(maxLink); i++)
            {
                linkList.Add((long.Parse(fixWildcardValue) + i + 1).ToString());
            }
            strDictionary[wildcard][fixWildcardValue] = linkList;
        }
        return fixWildcardValue;
    }

    private static string Mer(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string itemId,
        string funDy1
    )
    {
        // 尝试解析动态值
        if (long.TryParse(exportWildcardDyData[funDepends], out long value))
        {
            return (value + int.Parse(funDy1)).ToString();
        }

        // 如果解析失败，记录错误并返回一个默认值（例如0）或者抛出异常，具体根据业务需求
        Debug.Print($"Mer: 无法将 '{exportWildcardDyData[funDepends]}' 解析为 long 类型。");
        LogDisplay.RecordLine(
            "[{1}][{0}][{2}][{3}]",
            $"Mer: 无法将 '{exportWildcardDyData[funDepends]}' 解析为 long 类型。",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            funDepends,
            itemId
        );

        return exportWildcardDyData[funDepends]; // 或者 throw new FormatException($"无法将 '{exportWildcardDyData[funDepends]}' 解析为 long 类型。");
    }

    private static string MerB(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string itemId,
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

        if (int.TryParse(baseValue, out int baseValueTry))
        {
            if (int.TryParse(funDy1, out int funDy1Try))
            {
                if (int.TryParse(funDy2, out int funDy2Try))
                {
                    if (long.TryParse(exportWildcardDyData[funDepends], out long exValue))
                    {
                        if (int.TryParse(funDy3, out int funDy3Try))
                        {
                            if (baseValueTry + funDy1Try <= funDy2Try)
                            {
                                result = (exValue + funDy1Try).ToString();
                            }
                            else
                            {
                                result = (exValue + funDy1Try + funDy3Try).ToString();
                            }
                            return result;
                        }
                        Debug.Print($"MerB: 无法将 '{funDy3}' 解析为 int 类型。");
                        LogDisplay.RecordLine(
                            "[{1}][{0}][{2}][{3}]",
                            $"MerB: 无法将 '{funDy3}' 解析为 int 类型。",
                            DateTime.Now.ToString(CultureInfo.InvariantCulture),
                            funDy3,
                            itemId
                        );
                    }
                    Debug.Print($"MerB: 无法将 '{exportWildcardDyData[funDepends]}' 解析为 long 类型。");
                    LogDisplay.RecordLine(
                        "[{1}][{0}][{2}][{3}]",
                        $"MerB: 无法将 '{exportWildcardDyData[funDepends]}' 解析为 long 类型。",
                        DateTime.Now.ToString(CultureInfo.InvariantCulture),
                        funDepends,
                        itemId
                    );
                }
                Debug.Print($"MerB: 无法将 '{funDy2}' 解析为 int 类型。");
                LogDisplay.RecordLine(
                    "[{1}][{0}][{2}][{3}]",
                    $"MerB: 无法将 '{funDy2}' 解析为 int 类型。",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    funDy2,
                    itemId
                );
            }
            Debug.Print($"MerB: 无法将 '{funDy1}' 解析为 int 类型。");
            LogDisplay.RecordLine(
                "[{1}][{0}][{2}][{3}]",
                $"MerB: 无法将 '{funDy1}' 解析为 int 类型。",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                funDy1,
                itemId
            );
        }
        Debug.Print($"MerB: 无法将 '{baseValue}' 解析为 int 类型。");
        LogDisplay.RecordLine(
            "[{1}][{0}][{2}][{3}]",
            $"MerB: 无法将 '{baseValue}' 解析为 int 类型。",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            baseValue,
            itemId
        );
        return "0";
    }

    private static string MerTry(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3,
        List<string> idList
    )
    {
        string itemId = string.Empty;
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "3" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "10" : funDy3;
        string merB = MerB(exportWildcardDyData, funDepends, itemId, funDy1, funDy2, funDy3);
        string mer;
        mer = !idList.Contains(merB) ? Mer(exportWildcardDyData, funDepends, itemId, funDy1) : merB;
        return mer;
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
        if (funDy1 != "链类最大值")
        {
            return string.Empty;
        }
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
    //string funDy3
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "消耗量组" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "" : funDy2;
        //funDy3 = string.IsNullOrEmpty(funDy2) ? "" : funDy3;

        var funDy1Value = exportWildcardDyData[funDy1];
        var funDependsValue = exportWildcardDyData[funDepends];

        var funDy1ValueSplit = Regex.Split(funDy1Value, ",");

        var funDependsValueSplit = Regex.Split(funDependsValue, ",");

        //double numBit = 0;
        //if (double.TryParse(funDy3, out double intFunDy3))
        //{
        //    numBit = Math.Pow(10, intFunDy3 - 1);
        //}

        string result = "";
        if (funDy1ValueSplit.Length == funDependsValueSplit.Length)
        {
            for (int i = 0; i < funDy1ValueSplit.Length; i++)
            {
                string temp;
                if (funDy2 != "")
                {
                    ////寻找编号改为被寻找物品ID
                    //var funDy2Value = exportWildcardDyData[funDy2];
                    //if (long.TryParse(funDy2Value, out long funDy2ValueLong))
                    //{
                    //    temp =
                    //        $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]},{funDy2ValueLong + numBit * i}]";
                    //}
                    //else
                    //{
                    //    temp = $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]},{funDy2Value}]";
                    //}
                    temp =
                        $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]},{funDependsValueSplit[i]}]";
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
        var baseDicKey =
            exportWildcardDyData[funDy1]
                .Substring(0, exportWildcardDyData[funDy1].Length - int.Parse(funDy2)) + funDy3;
        var dependsDicValue = strDictionary[funDepends];

        var dependsValueList = dependsDicValue[baseDicKey];

        var baseNum = exportWildcardDyData[funDy1];

        if(dependsValueList.Contains(baseNum))
        {
            return string.Join(",", dependsValueList);
        }
        return String.Empty;
        
    }

    private static string GetDicKey(string funDepends)
    {
        //读取本地存储数据
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string filePath = Path.Combine(documentsPath, "strDic.csv");
        var fileDicData = LoadDictionaryFromFile(filePath);
        var dependsDicValue = fileDicData[funDepends];
        return string.Join(",", dependsDicValue.Keys);
    }

    private static string SplitArr(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "#" : funDy2;
        var dependsValue = exportWildcardDyData[funDepends];
        var dependsValueSplit = Regex.Split(dependsValue, funDy2);
        var result = dependsValueSplit[int.Parse(funDy1) - 1];
        return result;
    }

    private static string CollectRow(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3,
        string funDy4,
        string funDy5,
        Dictionary<string, List<string>> baseData,
        string id
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "消耗ID组" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "消耗量组" : funDy3;
        funDy4 = string.IsNullOrEmpty(funDy4) ? "20" : funDy4;
        funDy5 = string.IsNullOrEmpty(funDy5) ? "1" : funDy5;

        List<string> idList = baseData[id];
        List<string> funDy2List = baseData[funDy2];
        List<string> funDy3List = baseData[funDy3];

        var loopTimes = int.Parse(funDy4);
        if (long.TryParse(exportWildcardDyData[funDepends], out long collectRowId))
        {
            string strCollect = string.Empty;
            string stringSubCollect = string.Empty;
            string spawnCollect = string.Empty;

            // 首次的数据
            var idCollect = collectRowId.ToString();
            int findIndexFirst = idList.FindIndex(f => f == collectRowId.ToString());
            if (findIndexFirst != -1)
            {
                var funDy2Str = funDy2List[findIndexFirst];
                var funDy3Str = funDy3List[findIndexFirst];
                if (funDy2Str != String.Empty)
                {
                    var funDy2StrSplit = Regex.Split(funDy2Str, "#");
                    var funDy3StrSplit = Regex.Split(funDy3Str, "#");
                    if (funDy3StrSplit.Length == funDy2StrSplit.Length)
                    {
                        string temp;

                        for (int j = 0; j < funDy3StrSplit.Length; j++)
                        {
                            temp = $"[{funDy2StrSplit[j]},{funDy3StrSplit[j]},{funDy2StrSplit[j]}]";

                            strCollect += temp + ",";
                        }
                    }
                }
            }
            if (strCollect == String.Empty)
            {
                MessageBox.Show($"{idCollect}消耗数据为空，无法导出");
                return string.Empty;
            }

            strCollect = $"[{strCollect.Substring(0, strCollect.Length - 1)}]";

            // 其他次数据
            for (int i = 0; i < loopTimes; i++)
            {
                collectRowId += int.Parse(funDy1);
                int findIndex = idList.FindIndex(f => f == collectRowId.ToString());
                if (findIndex != -1)
                {
                    var funDy2Str = funDy2List[findIndex];
                    var funDy3Str = funDy3List[findIndex];
                    if (funDy2Str != String.Empty)
                    {
                        var funDy2StrSplit = Regex.Split(funDy2Str, "#");
                        var funDy3StrSplit = Regex.Split(funDy3Str, "#");
                        if (funDy3StrSplit.Length == funDy2StrSplit.Length)
                        {
                            string temp;

                            for (int j = 0; j < funDy3StrSplit.Length; j++)
                            {
                                temp =
                                    $"[{funDy2StrSplit[j]},{funDy3StrSplit[j]},{funDy2StrSplit[j]}]";
                                stringSubCollect += temp +",";
                            }
                            idCollect += "," + collectRowId;
                        }
                        else
                        {
                            spawnCollect = collectRowId.ToString();
                            break;
                        }
                    }
                    else
                    {
                        spawnCollect = collectRowId.ToString();
                        break;
                    }
                }
                else
                {
                    spawnCollect = collectRowId.ToString();
                    break;
                }
            }

            if (stringSubCollect == String.Empty)
            {
                strCollect = $"[{strCollect}]";
            }
            else
            {
                stringSubCollect = $"[{stringSubCollect.Substring(0, stringSubCollect.Length - 1)}]";
                strCollect = $"[{strCollect},{stringSubCollect}]";
            }

            idCollect = $"[{idCollect}]";
            if (funDy5 == "1")
            {
                return idCollect;
            }
            if (funDy5 == "2")
            {
                return strCollect;
            }
            if (funDy5 == "3")
            {
                return spawnCollect;
            }
        }
        Debug.Print($"MerB: 无法将 '{exportWildcardDyData[funDepends]}' 解析为 long 类型。");
        LogDisplay.RecordLine(
            "[{1}][{0}][{2}][{3}]",
            $"MerB: 无法将 '{exportWildcardDyData[funDepends]}' 解析为 long 类型。",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            exportWildcardDyData[funDepends],
            funDepends
        );
        return exportWildcardDyData[funDepends];
    }

    //获取动态值
    private static void GetDyWildcardValue(
        Dictionary<string, List<string>> baseData,
        Dictionary<string, string> exportWildcardDyData,
        string wildcard,
        string funDepends,
        int idCount
    )
    {
        var wildcardValuePattern = "#";
        // 获取有数据源的值
        if (funDepends.Contains("Var"))
        {
            var wildcardValueSplit = Regex.Split(funDepends, wildcardValuePattern);
            string fixWildcardValue = baseData[wildcardValueSplit[1]][idCount] ?? "";
            //ID组关键词替换
            if (wildcardValueSplit.Length == 3)
            {
                fixWildcardValue = fixWildcardValue.Replace(
                    wildcardValuePattern,
                    wildcardValueSplit[2]
                );
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
    private static Dictionary<string, Dictionary<string, List<string>>> LoadDictionaryFromFile(
        string filePath
    )
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
    #endregion

    #region LTE基础数据计算
    //去重复制
    public static void FilterRepeatValueCopy(CommandBarButton ctrl, ref bool cancelDefault)
    {
        //去重
        var mergedArray = FilterRepeatValue("", "", true);
        //复制
        PubMetToExcel.CopyArrayToClipboard(mergedArray);
    }

    //首次写入数据（指定范围内数据去重）
    public static void FirstCopyValue(CommandBarButton ctrl, ref bool cancelDefault)
    {
        object[,] copyArray = FilterRepeatValue(ActivityDataMinIndex, ActivityDataMaxIndex);

        var baseList = GetExcelListObjects("LTE【基础】", "LTE【基础】");
        if (baseList == null)
        {
            MessageBox.Show("LTE【基础】中的名称表-基础不存在");
            return;
        }
        var findList = GetExcelListObjects("LTE【寻找】", "LTE【寻找】");
        if (findList == null)
        {
            MessageBox.Show("LTE【寻找】中的名称表-寻找不存在");
            return;
        }
        //基础数据修改依赖数据
        var listObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#各类枚举"
        );
        var dataTypeArray = listObjectsDic["数据类型"];

        //基础数据整理
        var copyData = BaseData(copyArray, dataTypeArray);
        copyArray = copyData.fixArray;
        var errorTypeList = copyData.errorTypeList;
        if (errorTypeList.Count != 0)
        {
            //基础数据中存在错误类型
            var errorTypeListOnly = new HashSet<string>(errorTypeList);
            var errorStr = string.Join(",", errorTypeListOnly);
            MessageBox.Show($"基础数据中存在以下错误类型：{errorStr}");
        }
        ////基础List数据清理
        //baseList.DataBodyRange.ClearContents();
        ////基础List行数刷新
        //int newRowCount = copyArray.GetLength(0);
        //baseList.Resize(baseList.Range.Resize[newRowCount + 1, baseList.Range.Columns.Count]);
        //baseList.DataBodyRange.Value2 = copyArray;

        ////基础标记数据删除
        //var sheet = Wk.Worksheets["LTE【基础】"];
        //var oldTagRange = sheet.Range["A2:A10000"];
        //oldTagRange.Value2 = null;
        ////基础标记数据写入
        //var tagRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[copyArray.GetLength(0) + 1, 1]];
        //tagRange.Value2 = "+";

        //非Com写入数据,索引从0开始,效率确实更高,读取还是ListObject更方便
        var sheetName = "LTE【基础】";
        var rowMax = copyArray.GetLength(0);

        PubMetToExcel.WriteExcelDataC(sheetName, 1, 10000, BaseDataStartCol, BaseDataEndCol, null);
        PubMetToExcel.WriteExcelDataC(
            sheetName,
            1,
            rowMax,
            BaseDataStartCol,
            BaseDataEndCol,
            copyArray
        );

        baseList.Resize(baseList.Range.Resize[rowMax + 1, baseList.Range.Columns.Count]);

        PubMetToExcel.WriteExcelDataC(sheetName, 1, 10000, BaseDataTagCol, BaseDataTagCol, null);
        object[,] writeArray = new object[rowMax, 1];
        for (int i = 0; i < rowMax; i++)
            writeArray[i, 0] = "+";
        PubMetToExcel.WriteExcelDataC(
            sheetName,
            1,
            rowMax,
            BaseDataTagCol,
            BaseDataTagCol,
            writeArray
        );

        ////寻找数据整理
        var findArray = FindData(copyArray, dataTypeArray);
        ////寻找List数据清理
        //findList.DataBodyRange.ClearContents();
        ////寻找List行数刷新
        //int newFindRowCount = findArray.GetLength(0);
        //findList.Resize(baseList.Range.Resize[newFindRowCount + 1, findList.Range.Columns.Count]);
        //findList.DataBodyRange.Value2 = findArray;

        ////寻找标记数据删除
        //var sheetFind = Wk.Worksheets["LTE【寻找】"];
        //var oldTagFindRange = sheetFind.Range["A2:A10000"];
        //oldTagFindRange.Value2 = null;
        ////寻找标记数据写入
        //var tagFindRange = sheetFind.Range[
        //    sheetFind.Cells[2, 1],
        //    sheetFind.Cells[findArray.GetLength(0) + 1, 1]
        //];
        //tagFindRange.Value2 = "+";

        var sheetFindName = "LTE【寻找】";
        var rowFindMax = findArray.GetLength(0);

        PubMetToExcel.WriteExcelDataC(
            sheetFindName,
            1,
            10000,
            FindDataStartCol,
            FindDataEndCol,
            null
        );
        PubMetToExcel.WriteExcelDataC(
            sheetFindName,
            1,
            rowFindMax,
            FindDataStartCol,
            FindDataEndCol,
            findArray
        );

        findList.Resize(findList.Range.Resize[rowFindMax + 1, findList.Range.Columns.Count]);

        object[,] writeFindArray = new object[rowFindMax, 1];
        for (int i = 0; i < rowFindMax; i++)
            writeFindArray[i, 0] = "+";
        PubMetToExcel.WriteExcelDataC(
            sheetFindName,
            1,
            10000,
            FindDataTagCol,
            FindDataTagCol,
            null
        );
        PubMetToExcel.WriteExcelDataC(
            sheetFindName,
            1,
            rowFindMax,
            FindDataTagCol,
            FindDataTagCol,
            writeFindArray
        );
    }

    //更新写入数据（指定范围内数据去重），比对数据，更新数据状态
    public static void UpdateCopyValue(CommandBarButton ctrl, ref bool cancelDefault)
    {
        object[,] copyArray = FilterRepeatValue(ActivityDataMinIndex, ActivityDataMaxIndex);
        var list = GetExcelListObjects("LTE【基础】", "LTE【基础】");
        if (list == null)
        {
            MessageBox.Show("LTE【基础】中的名称表-基础不存在");
            return;
        }
        var findList = GetExcelListObjects("LTE【寻找】", "LTE【寻找】");
        if (findList == null)
        {
            MessageBox.Show("LTE【寻找】中的名称表-寻找不存在");
            return;
        }

        //基础数据修改依赖数据
        var listObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#各类枚举"
        );
        var dataTypeArray = listObjectsDic["数据类型"];

        //基础数据整理
        var copyData = BaseData(copyArray, dataTypeArray);
        copyArray = copyData.fixArray;
        var errorTypeList = copyData.errorTypeList;
        if (errorTypeList.Count != 0)
        {
            //基础数据中存在错误类型
            var errorTypeListOnly = new HashSet<string>(errorTypeList);
            var errorStr = string.Join(",", errorTypeListOnly);
            MessageBox.Show($"基础数据中存在以下错误类型：{errorStr}");
        }

        WriteDymaicData(copyArray, list, "LTE【基础】", 1, 33);

        //寻找数据整理
        var findArray = FindData(copyArray, dataTypeArray);
        WriteDymaicData(findArray, findList, "LTE【寻找】", 1, 9);
    }

    private static void WriteDymaicData(
        object[,] copyArray,
        ListObject list,
        string sheetName,
        int firstCol,
        int lastCol,
        int tagCol = 0
    )
    {
        //基础List数据清理
        object[,] oldListData = list.DataBodyRange.Value2;
        //基础数据和基础List数据对比
        var copyDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(copyArray);
        var oldListDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(oldListData);

        var tagDataGroup = TagData(copyDic, oldListDic);
        //基础数据对比后写入List
        object[,] data = PubMetToExcel.ConvertListToArray(tagDataGroup.Item1);

        var listRowCount = copyDic.Keys.Count;
        PubMetToExcel.WriteExcelDataC(sheetName, 1, listRowCount, firstCol, lastCol, null);
        int newRowCount = data.GetLength(0);
        PubMetToExcel.WriteExcelDataC(sheetName, 1, newRowCount, firstCol, lastCol, data);
        //整理原List大小
        list.Resize(list.Range.Resize[newRowCount + 1, list.Range.Columns.Count]);

        //基础数据对比标识写入Range
        PubMetToExcel.WriteExcelDataC(sheetName, 1, listRowCount, tagCol, tagCol, null);
        object[,] tagData = PubMetToExcel.ConvertList1ToArray(tagDataGroup.Item2);
        int tagRowCount = tagDataGroup.Item2.Count;
        PubMetToExcel.WriteExcelDataC(sheetName, 1, tagRowCount, tagCol, tagCol, tagData);
    }

    private static Tuple<List<List<string>>, List<string>> TagData(
        Dictionary<string, List<string>> copyDic,
        Dictionary<string, List<string>> oldListDic
    )
    {
        // 分类处理
        var added = new List<List<string>>();
        var deleted = new List<List<string>>();
        var modified = new List<List<string>>();
        var unchanged = new List<List<string>>();

        var addedTag = new List<string>();
        var deletedTag = new List<string>();
        var modifiedTag = new List<string>();
        var unchangedTag = new List<string>();

        // 新增项
        foreach (var key in copyDic.Keys.Where(k => !oldListDic.ContainsKey(k)))
        {
            added.Add(copyDic[key]);
            addedTag.Add("+");
        }
        // 删除项
        foreach (var key in oldListDic.Keys.Where(k => !copyDic.ContainsKey(k)))
        {
            deleted.Add(oldListDic[key]);
            deletedTag.Add("-");
        }
        // 修改/不变项
        foreach (var key in copyDic.Keys.Intersect(oldListDic.Keys))
        {
            var rowNew = copyDic[key];
            var rowOld = oldListDic[key];
            bool isModified = false;

            // 从第二列开始比较
            for (int i = 1; i < rowNew.Count; i++)
            {
                var newValue = rowNew[i];
                var oldValue = rowOld[i];
                if (newValue == null)
                {
                    newValue = "";
                }
                if (oldValue == null)
                {
                    oldValue = "";
                }

                if (newValue != oldValue)
                {
                    isModified = true;
                    break;
                }
            }

            if (isModified)
            {
                modified.Add(rowNew);
                modifiedTag.Add("*");
            }
            else
            {
                unchanged.Add(rowNew);
                unchangedTag.Add("#");
            }
        }
        //合并数据转数组
        List<List<string>> dataList = added
            .Concat(deleted)
            .Concat(modified)
            .Concat(unchanged)
            .ToList();
        List<string> tagList = addedTag
            .Concat(deletedTag)
            .Concat(modifiedTag)
            .Concat(unchangedTag)
            .ToList();
        return Tuple.Create(dataList, tagList);
    }

    //指定列范围的数据去重
    private static object[,] FilterRepeatValue(
        string min,
        string max,
        bool isSelect = false,
        bool isFilter = true
    )
    {
        var excel = NumDesAddIn.App;

        var sheet = excel.ActiveSheet as Worksheet;

        Range copyRange;
        if (!isSelect)
        {
            if (sheet == null)
            {
                MessageBox.Show("未找到表");
                return null;
            }
            var copyColMin = sheet.Range[min].Value2;
            var copyColMax = sheet.Range[max].Value2;
            copyRange = sheet.Range[sheet.Cells[3, copyColMin], sheet.Cells[10000, copyColMax]];
        }
        else
        {
            copyRange = excel.Selection;
        }
        if (copyRange == null)
        {
            // 如果没有选择任何内容，直接返回
            return null;
        }

        object[,] mergedArray;
        int index = 0;
        int baseIndex = 0;

        if (copyRange.Areas.Count > 1)
        {
            object[] areas = new object[copyRange.Areas.Count];

            // 获取每个区域的数据
            for (int i = 1; i <= copyRange.Areas.Count; i++)
            {
                areas[i - 1] = copyRange.Areas[i].Value2;
            }

            // 按列合并
            mergedArray = PubMetToExcel.MergeRanges(areas, false);
        }
        else
        {
            mergedArray = copyRange.Value2;
            index = 1;
            baseIndex = 1;
        }

        //去重
        if (isFilter)
        {
            mergedArray = PubMetToExcel.CleanRepeatValue(mergedArray, index, false, baseIndex);
        }
        return mergedArray;
    }

    //获取指定表的名称表
    private static ListObject GetExcelListObjects(string sheetName, string listName)
    {
        LogDisplay.RecordLine(
            "[{1}][{0}][{2}][{3}]",
            $"获取Excel ListObject: {sheetName} - {listName}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            sheetName,
            listName
        );
        var sheet = Wk.Worksheets[sheetName];
        // 获取ListObject并操作
        try
        {
            ListObject listObj = sheet.ListObjects[listName];
            return listObj;
        }
        catch (Exception e)
        {
            LogDisplay.RecordLine(
                    "[{1}][{0}][{2}][{3}]",
                    $"获取Excel ListObject: {sheetName} - {listName} 不存在",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    sheetName,
                    listName
                );
            throw;
        }
        
        
    }

    //原始数据改造
    private static (object[,] fixArray, List<string> errorTypeList) BaseData(
        object[,] baseArray,
        object[,] dataTypeArray
    )
    {
        var baseDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(baseArray);
        var dataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(dataTypeArray);

        var errorTypeList = new List<string>();

        foreach (var baseList in baseDic)
        {
            var key = baseList.Key;

            //寻找类型、寻找细类
            string findType;
            string findDetailType;

            string itemType = baseDic[key][6];

            //判断类型是否存在
            if (dataTypeDic.ContainsKey(itemType))
            {
                findType = dataTypeDic[itemType][2];
                findDetailType = dataTypeDic[itemType][3];
            }
            else
            {
                errorTypeList.Add(itemType);
                continue;
            }
            baseDic[key].Add(findType);
            baseDic[key].Add(findDetailType);

            //链长
            string linkMax = string.Empty;
            string currentName = baseDic[key][4];
            int countCurrent = baseDic
                .Values.Where(list => list.Count > 4)
                .Count(list => list[4] == currentName);

            if (itemType.Contains("链"))
            {
                if (countCurrent > 1)
                {
                    linkMax = countCurrent.ToString();
                }
            }
            baseDic[key].Add(linkMax);

            //五合提示
            string fiveMergeTip = string.Empty;
            string rank = baseDic[key][5];

            if (int.TryParse(rank, out int rankNum))
            {
                if (rankNum >= 6 && rankNum < countCurrent)
                {
                    fiveMergeTip = "35";
                }
            }
            baseDic[key].Add(fiveMergeTip);

            //消耗ID组、产出ID组、消耗量组、产出量组
            string consumeIdGroup;
            string productIdGroup;
            string consumeCountGroup;
            string productCountGroup;

            var consumeIdList = new List<string>();
            var productIdList = new List<string>();
            var consumeCountList = new List<string>();
            var productCountList = new List<string>();

            var idNameList = new List<int> { 9, 11, 13, 15, 17, 19, 21, 23 };
            var countNumList = new List<int> { 10, 12, 14, 16, 18, 20, 22, 24 };

            string firstPos = baseDic[key][1];
            var firstPosPre = firstPos.Split("-")[0];

            int onlyNum = 2;
            int num = 3;

            int countNum = 0;
            foreach (var idName in idNameList)
            {
                if (baseDic[key][idName] != null)
                {
                    var name = baseDic[key][idName];
                    if (name != string.Empty)
                    {
                        //先在唯一ID中查找
                        string matchId = baseDic
                            .FirstOrDefault(kv =>
                                kv.Value.Count > onlyNum && kv.Value[onlyNum] == firstPosPre + name
                            )
                            .Key;
                        if (matchId == null)
                        {
                            //后在ID中查找
                            matchId = baseDic
                                .FirstOrDefault(kv => kv.Value.Count > num && kv.Value[num] == name)
                                .Key;
                        }
                        if (matchId != null)
                        {
                            if (countNum < 4)
                            {
                                consumeIdList.Add(matchId);
                                consumeCountList.Add(baseDic[key][countNumList[countNum]]);
                            }
                            else
                            {
                                productIdList.Add(matchId);
                                productCountList.Add(baseDic[key][countNumList[countNum]]);
                            }
                        }
                    }
                }
                countNum++;
            }

            //如果没有消耗ID组则尝试查找谁消耗该ID
            //用来建立道具关系,此时组成的消耗组，只是为了建立寻找关系
            if (consumeIdList.Count == 0)
            {
                //该UD的代号和唯一代号
                var orginNum = baseDic[key][num];
                var orginOnlyNum = baseDic[key][onlyNum];
                string matchId = baseDic
                    .FirstOrDefault(kv =>
                        kv.Value.Count > 15
                        && kv.Value[1].Split("-")[0] + kv.Value[9] == orginOnlyNum
                    )
                    .Key;
                if (matchId == null)
                {
                    matchId = baseDic
                        .FirstOrDefault(kv =>
                            kv.Value.Count > 15
                            && kv.Value[1].Split("-")[0] + kv.Value[11] == orginOnlyNum
                        )
                        .Key;
                }
                if (matchId == null)
                {
                    matchId = baseDic
                        .FirstOrDefault(kv =>
                            kv.Value.Count > 15
                            && kv.Value[1].Split("-")[0] + kv.Value[13] == orginOnlyNum
                        )
                        .Key;
                }
                if (matchId == null)
                {
                    matchId = baseDic
                        .FirstOrDefault(kv =>
                            kv.Value.Count > 15
                            && kv.Value[1].Split("-")[0] + kv.Value[15] == orginOnlyNum
                        )
                        .Key;
                }
                //唯一代号找不到尝试寻找代号
                if (matchId == null)
                {
                    matchId = baseDic
                        .FirstOrDefault(kv => kv.Value.Count > 9 && kv.Value[9] == orginNum)
                        .Key;
                }
                if (matchId != null)
                {
                    consumeIdList.Add(matchId);
                    consumeCountList.Add("1");
                }
            }

            consumeIdGroup = string.Join("#", consumeIdList);
            productIdGroup = string.Join("#", productIdList);
            consumeCountGroup = string.Join("#", consumeCountList);
            productCountGroup = string.Join("#", productCountList);

            baseDic[key].Add(consumeIdGroup);
            baseDic[key].Add(productIdGroup);
            baseDic[key].Add(consumeCountGroup);
            baseDic[key].Add(productCountGroup);
        }
        var fixArray = PubMetToExcel.DictionaryTo2DArray(
            baseDic,
            baseDic.Count,
            baseDic[baseDic.Keys.First()].Count
        );

        return (fixArray, errorTypeList);
    }
    #endregion

    #region LTE寻找数据计算
    private static object[,] FindData(object[,] copyArray, object[,] dataTypeArray)
    {
        var findDic = new Dictionary<string, List<string>>();

        var copyDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(copyArray);

        var dataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(dataTypeArray);

        foreach (var key in copyDic.Keys)
        {
            var keyType = copyDic[key][6];
            if (!dataTypeDic.ContainsKey(keyType))
            {
                continue;
            }
            var dataType = dataTypeDic[keyType][4];
            if (dataType != "1")
            {
                continue;
            }
            //产出组
            var inputGroup = copyDic[key][29];
            var inputArray = inputGroup.Split("#");

            //正向查找
            for (int i = 0; i < inputArray.Length; i++)
            {
                //if (double.TryParse(key, out double intKey))
                //{
                //double findId = intKey + i * 100;
                var findTargetId = inputArray[i];
                var findId = findTargetId;

                if (findTargetId == string.Empty)
                {
                    continue;
                }
                var findTargetType = copyDic[findTargetId][25];
                var findTargetDetailType = copyDic[findTargetId][26];

                if (findTargetType != String.Empty)
                {
                    var findLinks = FindLinks(
                        findTargetDetailType,
                        findTargetType,
                        findTargetId,
                        copyDic,
                        out var findTips
                    );
                    if (findLinks != String.Empty)
                    {
                        var findIdStr = Convert.ToString(findId, CultureInfo.InvariantCulture);
                        if (!findDic.ContainsKey(findIdStr))
                        {
                            findDic.Add(findIdStr, new List<string>());
                        }
                        findDic[findIdStr].Add(findIdStr);
                        findDic[findIdStr].Add(copyDic[findIdStr][1]);
                        findDic[findIdStr].Add(copyDic[findIdStr][2]);
                        findDic[findIdStr].Add(copyDic[findIdStr][3]);
                        findDic[findIdStr].Add(copyDic[findIdStr][4]);
                        findDic[findIdStr].Add("寻-" + copyDic[key][6]);
                        findDic[findIdStr].Add(copyDic[findIdStr][7]);
                        findDic[findIdStr].Add(findTips);

                        var findLinksFix = findLinks.Substring(0, findLinks.Length - 1);
                        findLinksFix += ",{8,9999}";

                        findDic[findIdStr].Add(findLinksFix);
                    }
                }
                //}
            }

            //反向查找
            List<string> subMatchIDs = copyDic
                .Where(kv => kv.Value.Count > 31 && kv.Value[31].Contains(key))
                .Select(kv => kv.Key)
                .ToList();

            var subFindLinks = string.Empty;
            foreach (var findTargetId2 in subMatchIDs)
            {
                if (findTargetId2 != string.Empty)
                {
                    var subFindId = findTargetId2;
                    var findTargetType2 = copyDic[findTargetId2][25];
                    var findTargetDetailType2 = copyDic[findTargetId2][26];

                    if (findTargetType2 != string.Empty)
                    {
                        if (findTargetDetailType2 == string.Empty)
                        {
                            findTargetDetailType2 = "未找到细类";
                        }
                        if (findTargetType2 == "19")
                        {
                            subFindLinks +=
                                "{"
                                + findTargetType2
                                + ","
                                + findTargetDetailType2
                                + ","
                                + findTargetId2
                                + "},";
                        }
                        else if (findTargetType2 == "1")
                        {
                            subFindLinks += "{" + findTargetType2 + "," + findTargetId2 + "},";
                        }
                        else
                        {
                            subFindLinks += "{" + findTargetType2 + "," + findTargetId2 + "},";
                        }
                    }
                    if (subFindLinks != String.Empty)
                    {
                        if (!findDic.ContainsKey(subFindId))
                        {
                            findDic.Add(subFindId, new List<string>());
                            findDic[subFindId].Add(subFindId);
                            findDic[subFindId].Add(copyDic[subFindId][1]);
                            findDic[subFindId].Add(copyDic[subFindId][2]);
                            findDic[subFindId].Add(copyDic[subFindId][3]);
                            findDic[subFindId].Add(copyDic[subFindId][4]);
                            findDic[subFindId].Add("寻-" + copyDic[key][6]);
                            findDic[subFindId].Add(copyDic[subFindId][7]);
                            findDic[subFindId].Add("");

                            var findLinksFix = subFindLinks.Substring(0, subFindLinks.Length - 1);
                            findLinksFix += ",{8,9999}";

                            findDic[subFindId].Add(findLinksFix);
                        }
                    }
                }
            }
        }
        var findLinksArray = PubMetToExcel.DictionaryTo2DArray(
            findDic,
            findDic.Count,
            findDic[findDic.Keys.First()].Count
        );
        return findLinksArray;
    }

    private static string FindLinks(
        string findTargetDetailType,
        string findTargetType,
        string findTargetId,
        Dictionary<string, List<string>> copyDic,
        out string findTips
    )
    {
        var findLinks = string.Empty;
        findTips = string.Empty;
        //1层查找
        if (findTargetDetailType == string.Empty)
        {
            findTargetDetailType = "未找到细类";
        }
        if (findTargetType == "19")
        {
            findLinks +=
                "{" + findTargetType + "," + findTargetDetailType + "," + findTargetId + "},";
        }
        else if (findTargetType == "1")
        {
            findLinks += "{" + findTargetType + "," + findTargetId + "},";
        }
        else
        {
            findLinks += "{" + findTargetType + "," + findTargetId + "},";
        }
        //2层查找
        List<string> matchedIDs = copyDic
            .Where(kv => kv.Value.Count > 30 && kv.Value[30].Contains(findTargetId))
            .Select(kv => kv.Key)
            .ToList();
        //如果没有直接匹配的说明目标可能是链，需要继续查找
        if (matchedIDs.Count == 0)
        {
            findTargetId = findTargetId.Substring(0, findTargetId.Length - 2) + "01";
            matchedIDs = copyDic
                .Where(kv => kv.Value.Count > 30 && kv.Value[30].Contains(findTargetId))
                .Select(kv => kv.Key)
                .ToList();
        }
        if (matchedIDs.Count == 0)
        {
            if (double.TryParse(findTargetId, out _))
            {
                findTargetId = findTargetId.Substring(0, findTargetId.Length - 2) + "02";
                matchedIDs = copyDic
                    .Where(kv => kv.Value.Count > 30 && kv.Value[30].Contains(findTargetId))
                    .Select(kv => kv.Key)
                    .ToList();
            }
        }
        if (matchedIDs.Count == 0)
        {
            findTips = "{1,\"tip_obstacleItem\",2}";
        }
        else
        {
            int itemCount = 0;
            foreach (var findTargetId2 in matchedIDs)
            {
                if (findTargetId2 != string.Empty)
                {
                    var findTargetType2 = copyDic[findTargetId2][25];
                    var findTargetDetailType2 = copyDic[findTargetId2][26];

                    if (findTargetType2 != string.Empty)
                    {
                        if (findTargetDetailType2 == string.Empty)
                        {
                            findTargetDetailType2 = "未找到细类";
                        }
                        if (findTargetType2 == "19")
                        {
                            findLinks +=
                                "{"
                                + findTargetType2
                                + ","
                                + findTargetDetailType2
                                + ","
                                + findTargetId2
                                + "},";
                        }
                        else if (findTargetType2 == "1")
                        {
                            findLinks +=
                                "{" + findTargetType2 + "," + findTargetId2 + "},";
                        }
                        else
                        {
                            findLinks += "{" + findTargetType2 + "," + findTargetId2 + "},";
                        }
                        if (itemCount == 0)
                        {
                            if (findTargetDetailType == "4")
                            {
                                if (findTargetId is not null)
                                {
                                    findTips =
                                        "{3,"
                                        + findTargetId.Substring(0, findTargetId.Length - 2)
                                        + "00,"
                                        + findTargetId2
                                        + "}";
                                }
                            }
                            else
                            {
                                findTips = "{1,\"tip_obstacleItem\",1," + findTargetId2 + "}";
                            }
                        }
                    }
                }
                itemCount++;
            }
        }

        return findLinks;
    }

    #endregion

    #region LTE任务数据计算

    //首次写入数据
    public static void FirstCopyTaskValue(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sheetName = "LTE【任务】";
        var colIndexArray = PubMetToExcel.ReadExcelDataC(sheetName, 0, 0, 1, 1);
        double activtiyId = (double)colIndexArray[0, 0];

        object[,] copyTaskArray = FilterRepeatValue(
            ActivityDataMinIndex,
            ActivityDataMaxIndex,
            false,
            false
        );

        var taskList = GetExcelListObjects("LTE【任务】", "LTE【任务】");
        if (taskList == null)
        {
            MessageBox.Show("LTE【任务】中的名称表-任务不存在");
            return;
        }

        var baseList = GetExcelListObjects("LTE【基础】", "LTE【基础】");
        if (baseList == null)
        {
            MessageBox.Show("LTE【基础】中的名称表-基础不存在");
            return;
        }
        object[,] baseArray = baseList.DataBodyRange.Value2;

        //基础数据修改依赖数据
        var listObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#各类枚举"
        );
        object[,] dataTypeArray = listObjectsDic["任务类型"];

        //任务数据整理
        var copyTaskData = TaskData(copyTaskArray, dataTypeArray, baseArray, activtiyId);
        copyTaskArray = copyTaskData.taskArray;
        var errorTypeList = copyTaskData.errorTypeList;
        if (errorTypeList.Count != 0)
        {
            //基础数据中存在错误类型
            var errorTypeListOnly = new HashSet<string>(errorTypeList);
            var errorStr = string.Join(",", errorTypeListOnly);
            MessageBox.Show($"任务数据中存在以下错误类型：{errorStr}");
        }

        //非Com写入数据,索引从0开始,效率确实更高,读取还是ListObject更方便

        var rowMax = copyTaskArray.GetLength(0);

        PubMetToExcel.WriteExcelDataC(sheetName, 1, 10000, TaskDataStartCol, TaskDataEndCol, null);
        PubMetToExcel.WriteExcelDataC(
            sheetName,
            1,
            rowMax,
            TaskDataStartCol,
            TaskDataEndCol,
            copyTaskArray
        );

        taskList.Resize(taskList.Range.Resize[rowMax + 1, taskList.Range.Columns.Count]);

        PubMetToExcel.WriteExcelDataC(sheetName, 1, 10000, TaskDataTagCol, TaskDataTagCol, null);
        object[,] writeArray = new object[rowMax, 1];
        for (int i = 0; i < rowMax; i++)
            writeArray[i, 0] = "+";
        PubMetToExcel.WriteExcelDataC(
            sheetName,
            1,
            rowMax,
            TaskDataTagCol,
            TaskDataTagCol,
            writeArray
        );
    }

    //更新写入数据
    public static void UpdateCopyTaskValue(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sheetName = "LTE【任务】";
        var colIndexArray = PubMetToExcel.ReadExcelDataC(sheetName, 0, 0, 1, 1);
        double activtiyId = (double)colIndexArray[0, 0];

        object[,] copyTaskArray = FilterRepeatValue(
            ActivityDataMinIndex,
            ActivityDataMaxIndex,
            false,
            false
        );

        var taskList = GetExcelListObjects("LTE【任务】", "LTE【任务】");
        if (taskList == null)
        {
            MessageBox.Show("LTE【任务】中的名称表-任务不存在");
            return;
        }

        var baseList = GetExcelListObjects("LTE【基础】", "LTE【基础】");
        if (baseList == null)
        {
            MessageBox.Show("LTE【基础】中的名称表-基础不存在");
            return;
        }
        object[,] baseArray = baseList.DataBodyRange.Value2;

        //基础数据修改依赖数据
        var listObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#各类枚举"
        );
        object[,] dataTypeArray = listObjectsDic["任务类型"];

        //任务数据整理
        var copyTaskData = TaskData(copyTaskArray, dataTypeArray, baseArray, activtiyId);
        copyTaskArray = copyTaskData.taskArray;
        var errorTypeList = copyTaskData.errorTypeList;

        if (errorTypeList.Count != 0)
        {
            //基础数据中存在错误类型
            var errorTypeListOnly = new HashSet<string>(errorTypeList);
            var errorStr = string.Join(",", errorTypeListOnly);
            MessageBox.Show($"任务数据中存在以下错误类型：{errorStr}");
        }

        WriteDymaicData(copyTaskArray, taskList, "LTE【任务】", 14, 23, 13);
    }

    //原始数据改造
    private static (object[,] taskArray, List<string> errorTypeList) TaskData(
        object[,] copyTaskArray,
        object[,] taskDataTypeArray,
        object[,] baseList,
        double activtiyId
    )
    {
        var baseDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(baseList);
        var taskDataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(taskDataTypeArray);

        var taskTaskArrayCount = copyTaskArray.GetLength(0);
        var taskList = new List<List<string>>();

        var errorTypeList = new List<string>();

        for (int i = 1; i <= taskTaskArrayCount; i++)
        {
            var taskColDataList = new List<string>();
            var taskSubColDataList = new List<string>();

            var taskId = copyTaskArray[i, 1]?.ToString() ?? String.Empty;
            var taskDes = copyTaskArray[i, 2]?.ToString() ?? String.Empty;
            var taskTypeName = copyTaskArray[i, 3]?.ToString() ?? String.Empty;
            var taskDialogId = copyTaskArray[i, 4]?.ToString() ?? String.Empty;
            var taskTagetName = copyTaskArray[i, 5]?.ToString() ?? String.Empty;

            var taskSubId = copyTaskArray[i, 6]?.ToString() ?? String.Empty;
            var taskSubDes = copyTaskArray[i, 7]?.ToString() ?? String.Empty;
            var taskSubTypeName = copyTaskArray[i, 8]?.ToString() ?? String.Empty;
            var taskSubDialogId = copyTaskArray[i, 9]?.ToString() ?? String.Empty;
            var taskSubTagetName = copyTaskArray[i, 10]?.ToString() ?? String.Empty;

            //改造数据

            var fixMainData = FixTaskData(
                taskTypeName,
                taskDialogId,
                taskTagetName,
                activtiyId,
                taskDataTypeDic,
                baseDic
            );
            string taskTypeId = fixMainData[0];
            string taskTagetId = fixMainData[1];
            taskDialogId = fixMainData[2];

            var fixSubData = FixTaskData(
                taskSubTypeName,
                taskSubDialogId,
                taskSubTagetName,
                activtiyId,
                taskDataTypeDic,
                baseDic
            );
            string taskSubTypeId = fixSubData[0];
            string taskSubTagetId = fixSubData[1];
            taskSubDialogId = fixSubData[2];

            var taskNextId = string.Empty;

            //主线数据
            if (taskDes != string.Empty)
            {
                if (!taskDataTypeDic.ContainsKey(taskTypeName))
                {
                    errorTypeList.Add(taskTypeName);
                    continue;
                }
                taskColDataList.Add(taskId);
                taskColDataList.Add(taskDes);
                taskColDataList.Add(taskTypeName);
                taskColDataList.Add(taskTagetName);
                taskColDataList.Add(taskDialogId);
                taskColDataList.Add(taskTypeId);
                taskColDataList.Add(taskTagetId);

                //解锁下一个任务ID
                if (i != taskTaskArrayCount)
                {
                    if (double.TryParse(taskId, out double taskIdDouble))
                    {
                        taskIdDouble++;
                        if (taskSubId != string.Empty)
                        {
                            taskNextId = taskIdDouble + "," + taskSubId;
                        }
                        else
                        {
                            taskNextId = taskIdDouble.ToString(CultureInfo.InvariantCulture);
                        }
                    }
                }
                taskColDataList.Add(taskNextId);

                //目标所在地图
                if(taskTagetId == null)
                {
                    continue;
                }
                var taskTargetMapName = baseDic[taskTagetId][1];
                taskColDataList.Add(taskTargetMapName);

                //目标寻找关系
                var findTargetType = baseDic[taskTagetId][25];
                var findTargetDetailType = baseDic[taskTagetId][26];

                var findLinks = FindLinks(
                    findTargetDetailType,
                    findTargetType,
                    taskTagetId,
                    baseDic,
                    out _
                );

                taskTargetMapName = taskTargetMapName.Split("-")[0];
                var match = Regex.Match(taskTargetMapName, @"\d+");
                var taskTargetMapId = match.Success ? match.Value : "0";

                if (double.TryParse(taskTargetMapId, out double taskTargetMapIdDouble))
                {
                    taskTargetMapId = (taskTargetMapIdDouble + activtiyId).ToString(
                        CultureInfo.InvariantCulture
                    );
                }
                findLinks =
                    findLinks + "{20,\"UILteMapEntrance\"," + taskTargetMapId + "},{8,9999}";
                taskColDataList.Add(findLinks);
            }

            var taskSubNextId = string.Empty;
            //支线数据
            if (taskSubId != string.Empty)
            {
                if (!taskDataTypeDic.ContainsKey(taskSubTypeName))
                {
                    errorTypeList.Add(taskSubTypeName);
                    continue;
                }
                taskSubColDataList.Add(taskSubId);
                taskSubColDataList.Add(taskSubDes);
                taskSubColDataList.Add(taskSubTypeName);
                taskSubColDataList.Add(taskSubTagetName);
                taskSubColDataList.Add(taskSubDialogId);
                taskSubColDataList.Add(taskSubTypeId);
                taskSubColDataList.Add(taskSubTagetId);

                //解锁下一个任务ID
                if (i != taskTaskArrayCount)
                {
                    taskSubNextId = copyTaskArray[i + 1, 6]?.ToString() ?? string.Empty;
                }
                taskSubColDataList.Add(taskSubNextId);

                //目标所在地图
                var taskSubTargetMapName = baseDic[taskSubTagetId][1];
                taskSubColDataList.Add(taskSubTargetMapName);

                //目标寻找关系
                var findSubTargetType = baseDic[taskSubTagetId][25];
                var findSubTargetDetailType = baseDic[taskSubTagetId][26];

                var findSubLinks = FindLinks(
                    findSubTargetDetailType,
                    findSubTargetType,
                    taskSubTagetId,
                    baseDic,
                    out _
                );

                taskSubTargetMapName = taskSubTargetMapName.Split("-")[0];
                var match = Regex.Match(taskSubTargetMapName, @"\d+");
                var taskSubTargetMapId = match.Success ? match.Value : "0";

                if (double.TryParse(taskSubTargetMapId, out double taskSubTargetMapIdDouble))
                {
                    taskSubTargetMapId = (taskSubTargetMapIdDouble + activtiyId).ToString(
                        CultureInfo.InvariantCulture
                    );
                }
                findSubLinks =
                    findSubLinks + "{20,\"UILteMapEntrance\"," + taskSubTargetMapId + "},{8,9999}";
                taskSubColDataList.Add(findSubLinks);
            }

            if (taskColDataList.Count != 0)
            {
                taskList.Add(taskColDataList);
            }
            if (taskSubColDataList.Count != 0)
            {
                taskList.Add(taskSubColDataList);
            }
        }
        var taskArray = PubMetToExcel.ConvertListToArray(taskList);
        return (taskArray, errorTypeList);
    }

    private static List<string> FixTaskData(
        string taskTypeName,
        string taskDialogId,
        string taskTagetName,
        double activtiyId,
        Dictionary<string, List<string>> taskDataTypeDic,
        Dictionary<string, List<string>> baseDic
    )
    {
        var fixData = new List<string>();

        string taskTypeId = string.Empty;
        string taskTagetId = string.Empty;

        if (taskTypeName != string.Empty)
        {
            taskTypeId = taskDataTypeDic[taskTypeName][1] ?? string.Empty;
        }
        if (taskTagetName != string.Empty)
        {
            taskTagetId = baseDic
                .FirstOrDefault(kv => kv.Value.Count > 2 && kv.Value[2] == taskTagetName)
                .Key;
        }
        if (taskDialogId != string.Empty)
        {
            if (double.TryParse(taskDialogId, out double taskTagetIdDouble))
            {
                taskDialogId = Convert.ToString(
                    taskTagetIdDouble + activtiyId,
                    CultureInfo.InvariantCulture
                );
            }
        }

        fixData.Add(taskTypeId);
        if(taskTagetId ==null)
        {
            MessageBox.Show($"任务目标ID{taskTagetName}不存在");
        }
        fixData.Add(taskTagetId);
        fixData.Add(taskDialogId);

        return fixData;
    }
    #endregion
}
