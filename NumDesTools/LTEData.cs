using System.Runtime.Versioning;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Documents;
using Microsoft.VisualBasic;
using NumDesTools.UI;
using OfficeOpenXml;

namespace NumDesTools;

[SupportedOSPlatform("windows")]
public class LteData
{
    // Introduce optional Excel host for easier testing; fall back to NumDesAddIn if not provided
    public static IExcelHost ExcelHostInstance { get; set; }

    private static Workbook Wk =>
        (ExcelHostInstance?.GetActiveWorkbook() as Workbook) ?? AppServices.App.ActiveWorkbook;

    private static string WkPath => Wk.Path;

    private static readonly Regex _digitsRegex = new(@"\d+", RegexOptions.Compiled);

    private static readonly Dictionary<string, (string id, string idType)> _sheetTypeMap = new(
        StringComparer.Ordinal
    )
    {
        ["LTE【基础】"] = ("数据编号", "类型"),
        ["LTE【任务】"] = ("任务编号", "类型"),
        ["LTE【寻找】"] = ("寻找编号", "类型"),
        ["LTE【地组】"] = ("地组编号", "类型"),
        ["LTE【通用】"] = ("数据编号", "类型"),
    };

    private const int BaseDataTagCol = 0;
    private const int BaseDataStartCol = 1;
    private const int FindDataTagCol = 0;
    private const int FindDataStartCol = 1;
    private const int FindDataEndCol = 9;
    private const int TaskDataTagCol = 15;
    private const int TaskDataStartCol = 16;
    private const int TaskDataEndCol = 27;
    private const int FieldDataTagCol = 13;
    private const int FieldDataStartCol = 14;
    private const int FieldDataEndCol = 23;

    // baseDic 原始列索引（与 LTE【基础】表列顺序对应）
    private const int BaseDicColPrefabId = 1;
    private const int BaseDicColIconId = 2;
    private const int BaseDicColFirstMap = 3;
    private const int BaseDicColOnlyName = 4;
    private const int BaseDicColName = 5;
    private const int BaseDicColPackage = 6;
    private const int BaseDicColLevel = 7;
    private const int BaseDicColType = 8;

    // baseDic 计算列索引（由 BaseData() 按顺序 .Add() 追加，原始列之后）
    private const int BaseDicCalcFindType = 9;
    private const int BaseDicCalcFindDetailType = 10;
    private const int BaseDicCalcLinkMax = 11;
    private const int BaseDicCalcFiveMergeTip = 12;

    // WriteExcelDataC 的"清空到末行"哨兵值
    private const int ClearToLastRow = 10000;

    // 活动编号除数（设计表 B1 存储的是 activityId * 10000）
    private const int ActivityIdDivisor = 10000;

    // LTE 任务/地组 寻找关系格式串片段
    private const string FindLinkMapEntrance = "{20,\"UILteMapEntrance\",";
    private const string FindLinkEndNormal = "},{8,9999}";
    private const string FindLinkEndPortal = "},{8,9993}";

    private static readonly List<object> _taskTitleArray = new List<object>()
    {
        "任务编号",
        "任务描述",
        "类型",
        "任务目标",
        "触发对话",
        "类型ID",
        "任务目标ID",
        "解锁任务",
        "所在地图",
        "寻找关系",
        "任务时间",
        "目标等级",
    };

    private const string ActivityIdIndex = "B1";
    private const string ActivityDataMinIndex = "C1";
    private const string ActivityDataMaxIndex = "D1";
    private const string ActivityNameMinIndex = "E1";
    private const string ActivityFieldIndex = "G1";

    private static Dictionary<string, string> OutputWildcardPubDic
    {
        get
        {
            var sheet = Wk.Worksheets["LTE【设计】"];
            return new Dictionary<string, string>
            {
                ["活动编号"] = (
                    sheet.Range[ActivityIdIndex].Value2 / ActivityIdDivisor
                )?.ToString(),
                ["活动备注"] = sheet.Range[ActivityNameMinIndex].Value2?.ToString(),
                ["活动地组"] = sheet.Range[ActivityFieldIndex].Value2?.ToString(),
            };
        }
    }
    public static (string Name, string Email) GitConfig = SvnGitTools.GetGitUserInfo();

    private static void RunLteCommand(
        CommandBarButton ctrl,
        ref bool cancelDefault,
        string commandName,
        System.Action body
    )
    {
        cancelDefault = true;
        try
        {
            body();
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[LTEData] {commandName} CRASH: {ex}");
            MessageBox.Show(
                $"操作失败，已记录日志。\n{ex.Message}",
                "错误",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    // 读取当前工作簿中指定 sheet 的 ListObject；找不到时弹提示并返回 null
    private static ListObject RequireListObject(string sheetName, string tableName)
    {
        var list = PubMetToExcel.GetExcelListObjects(sheetName, tableName);
        if (list == null)
            MessageBox.Show($"「{sheetName}」中找不到名称为「{tableName}」的表格，请检查");
        return list;
    }

    #region LTE数据配置导出
    //导出LTE数据配置
    public static void ExportLteDataConfigFirst(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(ExportLteDataConfigFirst),
            () => ExportLteDataConfig(true, GitConfig.Name)
        );

    public static void ExportLteDataConfigUpdate(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(ExportLteDataConfigUpdate),
            () => ExportLteDataConfig(false, GitConfig.Name)
        );

    public static void ExportLteDataConfig(bool isFirst, string userName)
    {
        //Epplus获取[LTE配置【导出】]表的ListObject
        var sheet = Wk.ActiveSheet;
        var sheetName = sheet.Name;
        var outputData = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "LTE配置【导出】"
        );

        // 自动匹配不同用户名的配置
        if (userName == null)
        {
            userName = String.Empty;
        }
        var listName = $"{sheetName}_通配符{userName}";
        if (!outputData.ContainsKey(listName))
        {
            var choose1 = MessageBox.Show(
                $"配置表没有包含【{userName}】的配置，是否使用默认用户配置Yes",
                "确认",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (choose1 == DialogResult.Yes)
            {
                listName = $"{sheetName}_通配符";
            }
            else
            {
                var selfInputName = Interaction.InputBox("输入使用的配置用户名", "自定义用户名");

                listName = $"{sheetName}_通配符{selfInputName}";

                if (!outputData.ContainsKey(listName))
                {
                    var choose2 = MessageBox.Show(
                        $"输入的用户名【{selfInputName}】配置不存在，使用默认配置Yes,终止导出操作No",
                        "确认",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );
                    if (choose2 == DialogResult.Yes)
                    {
                        listName = $"{sheetName}_通配符";
                    }
                    else
                    {
                        return;
                    }
                }
            }
        }

        var outputWildcardArray = outputData[listName];
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

        if (_sheetTypeMap.ContainsKey(sheetName))
        {
            var kv = _sheetTypeMap[sheetName];
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
                sheet.Cells[isFirstTagStartRow, isFirstTagStartCol - 1],
                sheet.Cells[isFirstTagEndRow, isFirstTagStartCol - 1]
            ];

            object[,] firstTagArray = firstTagRange.Value2;

            // 合并数据
            var mergeBaseArray = PubMetToExcel.Merge2DArrays1(sheetBaseArray, firstTagArray);

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
                PluginLog.Write("[LteData][GetModelValue] 表格不存在");
            }
        }
        modelExcel?.Dispose();
        return false;
    }

    //个别导出LTE数据配置
    private sealed record RowWriteContext(
        ExcelWorksheet TargetSheet,
        ExcelPackage TargetExcel,
        int WriteCol,
        string[] ColTitles,
        string[] ColTypes,
        KeyValuePair<string, Dictionary<(object, object), string>> ModelSheet,
        Dictionary<string, string> ExportWildcardData,
        Dictionary<string, string> ExportWildcardDyData,
        Dictionary<string, Dictionary<string, List<string>>> StrDictionary,
        Dictionary<string, List<string>> BaseData,
        string Id,
        HashSet<string> DataRepeatWritten,
        List<(string, int, int, string, string, string)> CheckResult
    );

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

        // 检查Excel单元格值是否非法
        var checkResult = new List<(string, int, int, string, string, string)>();

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
                                status,
                            }
                    )
                    .Where(x => x.status?.ToString() is "+" or "-" or "*")
                    .ToList();
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
                    $"[{DateTime.Now}] , {modelSheetName}【#LTE数据模版】中创建的文件名不存在"
                );
            }

            if (targetSheet != null)
            {
                AppServices.App.StatusBar = $"导出：{modelSheetName}";

                var writeCol = targetSheet.Dimension.End.Column;

                // 预读标题行和类型行，避免内层循环重复读 Cells
                var colTitles = new string[writeCol + 1];
                var colTypes = new string[writeCol + 1];
                for (int j = 2; j <= writeCol; j++)
                {
                    colTitles[j] = targetSheet.Cells[2, j].Value?.ToString() ?? "";
                    colTypes[j] = targetSheet.Cells[3, j].Value?.ToString() ?? "";
                }

                // 预建 ID→行号索引，消灭 FindSourceRow 的 O(n) 循环
                var idRowIndex = new Dictionary<string, int>(StringComparer.Ordinal);
                var endRow = targetSheet.Dimension.End.Row;
                for (int r = 2; r <= endRow; r++)
                {
                    var v = targetSheet.Cells[r, 2].Value?.ToString() ?? "";
                    if (v != "")
                        idRowIndex[v] = r;
                }

                var exportWildcardDyData = new Dictionary<string, string>(exportWildcardData);
                bool dataWritten = false;
                var dataRepeatWritten = new HashSet<string>();
                var modelSheetTypes = new HashSet<object>(
                    modelSheet.Value.Keys.Select(k => k.Item1)
                );
                var rowCtx = new RowWriteContext(
                    targetSheet,
                    targetExcel,
                    writeCol,
                    colTitles,
                    colTypes,
                    modelSheet,
                    exportWildcardData,
                    exportWildcardDyData,
                    strDictionary,
                    baseData,
                    id,
                    dataRepeatWritten,
                    checkResult
                );

                // 阶段一：收集操作意图，不修改表
                var deleteRows = new List<int>(); // 要删除的行号
                var updateOps = new List<(int rowIndex, int idCount)>(); // 原地覆写
                var insertOps = new List<(string itemId, int idCount)>(); // 新增（insert after 相似ID）
                var appendOps = new List<int>(); // 追加到末尾（无相似ID）

                if (dataStatusListNew != null)
                {
                    for (int idCount = 0; idCount < idList.Count; idCount++)
                    {
                        string itemId = idList[idCount] ?? "";
                        if (itemId == "")
                            continue;
                        string itemType = typeList[idCount] ?? "";
                        if (!modelSheetTypes.Contains(itemType))
                            continue;

                        string status = dataStatusList[idCount] ?? "";
                        if (status is not ("+" or "-" or "*"))
                            continue;

                        // 用构造 ID 查索引：帮助表等场景 id 列模版展开后与原表 itemId 不同
                        foreach (var wildcardDy in exportWildcardData)
                            GetDyWildcardValue(
                                baseData,
                                exportWildcardDyData,
                                wildcardDy.Key,
                                wildcardDy.Value,
                                idCount
                            );
                        var idColModelValue =
                            colTitles[2] != ""
                            && modelSheet.Value.TryGetValue(
                                (itemType, (object)colTitles[2]),
                                out var idModel
                            )
                                ? idModel
                                : null;
                        var lookupId =
                            idColModelValue != null
                                ? AnalyzeWildcard(
                                    idColModelValue,
                                    exportWildcardData,
                                    exportWildcardDyData,
                                    strDictionary,
                                    baseData,
                                    id,
                                    itemId
                                )
                                : itemId;
                        if (string.IsNullOrEmpty(lookupId))
                            lookupId = itemId;

                        idRowIndex.TryGetValue(lookupId, out int existingRow);

                        if (existingRow != 0)
                        {
                            if (status == "-")
                                deleteRows.Add(existingRow);
                            else if (status == "*")
                                updateOps.Add((existingRow, idCount));
                            // status == "+" 且已存在：跳过（已写入过）
                        }
                        else
                        {
                            if (status == "-")
                                continue; // 目标表不存在，跳过
                            // "+" 或 "*"：需要新增
                            if (itemId.Length >= 6)
                                insertOps.Add((itemId, idCount));
                            else
                                appendOps.Add(idCount);
                        }
                    }
                }
                else
                {
                    // 首次模式：全量追加
                    for (int idCount = 0; idCount < idList.Count; idCount++)
                    {
                        string itemId = idList[idCount] ?? "";
                        if (itemId == "")
                            continue;
                        if (!modelSheetTypes.Contains(typeList[idCount] ?? ""))
                            continue;
                        appendOps.Add(idCount);
                    }
                }

                // 阶段二执行：先删（倒序），再更新，最后插入/追加

                // 2a. 删除：倒序，避免行号偏移
                deleteRows.Sort((a, b) => b.CompareTo(a));
                foreach (var rowToDel in deleteRows)
                {
                    targetSheet.DeleteRow(rowToDel);
                    // 修正 updateOps 中受影响的行号
                    for (int i = 0; i < updateOps.Count; i++)
                    {
                        if (updateOps[i].rowIndex > rowToDel)
                            updateOps[i] = (updateOps[i].rowIndex - 1, updateOps[i].idCount);
                    }
                    idRowIndex = idRowIndex
                        .Where(kv => kv.Value != rowToDel)
                        .ToDictionary(
                            kv => kv.Key,
                            kv => kv.Value > rowToDel ? kv.Value - 1 : kv.Value,
                            StringComparer.Ordinal
                        );
                    dataWritten = true;
                }

                // 2b. 原地更新
                foreach (var (rowIndex, idCount) in updateOps)
                {
                    string itemId = idList[idCount] ?? "";
                    string itemType = typeList[idCount] ?? "";

                    foreach (var wildcardDy in exportWildcardData)
                        GetDyWildcardValue(
                            baseData,
                            exportWildcardDyData,
                            wildcardDy.Key,
                            wildcardDy.Value,
                            idCount
                        );

                    bool rowChanged = false;
                    var rowWriteData = new List<(int col, string val, string colType)>();
                    for (int j = 2; j <= writeCol; j++)
                    {
                        var cellTitle = colTitles[j];
                        if (cellTitle == "")
                            continue;
                        if (
                            !modelSheet.Value.TryGetValue(
                                (itemType, (object)cellTitle),
                                out var cellModelValue
                            )
                        )
                            continue;

                        var cellRealValue = AnalyzeWildcard(
                            cellModelValue,
                            exportWildcardData,
                            exportWildcardDyData,
                            strDictionary,
                            baseData,
                            id,
                            itemId
                        );

                        if (j == 2 && cellRealValue == string.Empty)
                            break;
                        if (j == 2 && dataRepeatWritten.Contains(cellRealValue))
                            break;
                        if (j == 2)
                            dataRepeatWritten.Add(cellRealValue);

                        var existingVal = targetSheet.Cells[rowIndex, j].Value?.ToString() ?? "";
                        if (existingVal != cellRealValue)
                        {
                            rowWriteData.Add((j, cellRealValue, colTypes[j]));
                            rowChanged = true;
                        }
                    }

                    if (!rowChanged)
                        continue;

                    foreach (var (col, val, colType) in rowWriteData)
                    {
                        var newErrors = PubMetToExcel.ExcelCellValueFormatCheck(
                            val,
                            colType,
                            targetSheet.Name,
                            targetExcel.File.FullName,
                            rowIndex - 1,
                            col - 1
                        );
                        foreach (var e in newErrors)
                            PluginLog.Write(
                                $"[LTEData][格式错误] itemId={itemId} {e.Item4} R{e.Item2}C{e.Item3}: {e.Item5} = {e.Item1}"
                            );
                        checkResult.AddRange(
                            newErrors.Select(e =>
                                (
                                    $"itemId={itemId} {e.Item1}",
                                    e.Item2,
                                    e.Item3,
                                    e.Item4,
                                    e.Item5,
                                    e.Item6
                                )
                            )
                        );
                        // val 来自 AnalyzeWildcard，永远是 string；本该是数字的字段直接写会
                        // 落成 sharedStrings 引用，纯数字大多不重复，去重零收益，纯体积浪费。
                        CellValueNormalizer.ApplyTo(targetSheet.Cells[rowIndex, col], val);
                    }
                    dataWritten = true;
                }

                // 2c. 新增（insert after 相似ID）
                // 插入时行号会变，记录累计偏移
                var insertsSorted = insertOps
                    .Select(op =>
                    {
                        var activeId = op.itemId.Substring(0, 6);
                        var regex = new Regex($"^{activeId}\\d{{4}}$");
                        int baseRow = PubMetToExcel.FindSourceRowBlur(targetSheet, 2, regex);
                        return (baseRow, op.itemId, op.idCount);
                    })
                    .Where(x => x.baseRow != -1)
                    .OrderBy(x => x.baseRow)
                    .ToList();

                // 无相似ID的追加到末尾
                var tailInserts = insertOps
                    .Select(op =>
                    {
                        var activeId = op.itemId.Substring(0, 6);
                        var regex = new Regex($"^{activeId}\\d{{4}}$");
                        int baseRow = PubMetToExcel.FindSourceRowBlur(targetSheet, 2, regex);
                        return (baseRow, op.idCount);
                    })
                    .Where(x => x.baseRow == -1)
                    .Select(x => x.idCount)
                    .Concat(appendOps)
                    .ToList();

                int rowOffset = 0;
                foreach (var (baseRow, itemId, idCount) in insertsSorted)
                {
                    string itemType = typeList[idCount] ?? "";
                    int writeRow = baseRow + 1 + rowOffset;
                    targetSheet.InsertRow(writeRow, 1);
                    rowOffset++;

                    foreach (var wildcardDy in exportWildcardData)
                        GetDyWildcardValue(
                            baseData,
                            exportWildcardDyData,
                            wildcardDy.Key,
                            wildcardDy.Value,
                            idCount
                        );

                    bool wrote = WriteRowData(
                        rowCtx,
                        writeRow,
                        itemId,
                        itemType,
                        idCount,
                        isFirst: true
                    );
                    if (wrote)
                        dataWritten = true;
                }

                foreach (var idCount in tailInserts)
                {
                    string itemId = idList[idCount] ?? "";
                    string itemType = typeList[idCount] ?? "";
                    int writeRow = targetSheet.Dimension.End.Row + 1;

                    foreach (var wildcardDy in exportWildcardData)
                        GetDyWildcardValue(
                            baseData,
                            exportWildcardDyData,
                            wildcardDy.Key,
                            wildcardDy.Value,
                            idCount
                        );

                    bool wrote = WriteRowData(
                        rowCtx,
                        writeRow,
                        itemId,
                        itemType,
                        idCount,
                        isFirst: true
                    );
                    if (wrote)
                        dataWritten = true;
                }

                if (dataWritten)
                    FileLockHelper.SaveWithRetry(targetExcel, targetExcel.File.FullName);
            }

            targetExcel?.Dispose();
        }
        // 展示Excel单元格数据格式错误
        if (checkResult.Count > 0)
        {
            var ctpCheckValueName = "单元格数据格式检查";
            NumDesCTP.DeleteCTP(true, ctpCheckValueName);
            _ = (SheetCellSeachResult)
                NumDesCTP.ShowCTP(
                    800,
                    ctpCheckValueName,
                    true,
                    ctpCheckValueName,
                    new SheetCellSeachResult(checkResult),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }

        // 输出字典数据
        if (strDictionary.Count > 0)
        {
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string filePath = Path.Combine(documentsPath, "strDic.csv");
            SaveDictionaryToFile(strDictionary, filePath);
        }
    }

    private static bool WriteRowData(
        RowWriteContext ctx,
        int writeRow,
        string itemId,
        string itemType,
        int idCount,
        bool isFirst
    )
    {
        bool wrote = false;
        for (int j = 2; j <= ctx.WriteCol; j++)
        {
            var cellTitle = ctx.ColTitles[j];
            if (cellTitle == "")
                continue;
            if (
                !ctx.ModelSheet.Value.TryGetValue(
                    (itemType, (object)cellTitle),
                    out var cellModelValue
                )
            )
                continue;

            var cellRealValue = AnalyzeWildcard(
                cellModelValue,
                ctx.ExportWildcardData,
                ctx.ExportWildcardDyData,
                ctx.StrDictionary,
                ctx.BaseData,
                ctx.Id,
                itemId
            );

            if (j == 2 && cellRealValue == string.Empty)
                break;
            if (j == 2 && ctx.DataRepeatWritten.Contains(cellRealValue))
                break;
            if (j == 2)
                ctx.DataRepeatWritten.Add(cellRealValue);

            if (!isFirst && ctx.TargetSheet.Cells[writeRow, j].Value?.ToString() == cellRealValue)
                continue;

            var newErrors = PubMetToExcel.ExcelCellValueFormatCheck(
                cellRealValue,
                ctx.ColTypes[j],
                ctx.TargetSheet.Name,
                ctx.TargetExcel.File.FullName,
                writeRow - 1,
                j - 1
            );
            foreach (var e in newErrors)
                PluginLog.Write(
                    $"[LTEData][格式错误] itemId={itemId} {e.Item4} R{e.Item2}C{e.Item3}: {e.Item5} = {e.Item1}"
                );
            ctx.CheckResult.AddRange(
                newErrors.Select(e =>
                    ($"itemId={itemId} {e.Item1}", e.Item2, e.Item3, e.Item4, e.Item5, e.Item6)
                )
            );
            CellValueNormalizer.ApplyTo(ctx.TargetSheet.Cells[writeRow, j], cellRealValue);
            wrote = true;
        }
        return wrote;
    }

    // delegate pure-logic helper implementations to LteCore to centralize logic and enable testing
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
        // delegate to LteCore for pure logic
        return LteCore.AnalyzeWildcard(
            cellModelValue,
            exportWildcardData,
            exportWildcardDyData,
            strDictionary,
            baseData,
            id,
            itemId
        );
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
        LteCore.GetDyWildcardValue(baseData, exportWildcardDyData, wildcard, funDepends, idCount);
    }

    //自定义字典初始化
    internal static void InitializeDictionary(
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        string key,
        string subKey
    )
    {
        LteCore.InitializeDictionary(strDictionary, key, subKey);
    }

    //循环数字
    internal static List<int> LoopNumber(int start, int max)
    {
        return LteCore.LoopNumber(start, max);
    }

    //strDic输出到文件
    private static void SaveDictionaryToFile(
        Dictionary<string, Dictionary<string, List<string>>> dictionary,
        string filePath
    )
    {
        LteCore.SaveDictionaryToFile(dictionary, filePath);
    }

    #endregion

    #region LTE基础数据计算
    //去重复制
    public static void FilterRepeatValueCopy(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(FilterRepeatValueCopy),
            () =>
            {
                var mergedArray = FilterRepeatValue("", "", true);
                PubMetToExcel.CopyArrayToClipboard(mergedArray);
            }
        );

    //首次写入数据（指定范围内数据去重）
    public static void FirstCopyValue(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(FirstCopyValue),
            () =>
            {
                object[,] copyArray = FilterRepeatValue(ActivityDataMinIndex, ActivityDataMaxIndex);
                object[,] copyTilteArray = ColTitleValue(
                    ActivityDataMinIndex,
                    ActivityDataMaxIndex
                );

                var baseList = RequireListObject("LTE【基础】", "LTE【基础】");
                if (baseList == null)
                    return;
                var findList = RequireListObject("LTE【寻找】", "LTE【寻找】");
                if (findList == null)
                    return;
                //基础数据修改依赖数据
                var listObjectsDic = PubMetToExcel.GetExcelListObjects(
                    WkPath,
                    "#【A-LTE】数值大纲.xlsx",
                    "#各类枚举"
                );
                var dataTypeArray = GetTableData(listObjectsDic, "#各类枚举", "数据类型");
                if (dataTypeArray == null)
                    return;

                //基础数据整理
                var copyData = BaseData(copyArray, dataTypeArray, copyTilteArray);
                copyArray = copyData.fixArray;
                copyTilteArray = copyData.fixTitleArray;
                var colCount = copyTilteArray.Length;
                var errorTypeList = copyData.errorTypeList;

                if (errorTypeList.Count != 0)
                {
                    var emptyKeyErrors = errorTypeList.Where(e => e.StartsWith('[')).ToList();
                    var unknownTypeErrors = errorTypeList.Where(e => !e.StartsWith('[')).ToList();
                    var sb = new StringBuilder("基础数据存在以下问题，请修复后重试：\n");
                    if (emptyKeyErrors.Count > 0)
                        sb.AppendLine(
                            $"\n【类型字段为空】共 {emptyKeyErrors.Count} 行：\n"
                                + string.Join("\n", emptyKeyErrors)
                        );
                    if (unknownTypeErrors.Count > 0)
                    {
                        var unknownOnly = new HashSet<string>(unknownTypeErrors);
                        sb.AppendLine($"\n【类型不在枚举中】{string.Join(", ", unknownOnly)}");
                    }
                    MessageBox.Show(sb.ToString());
                    return;
                }
                ////基础List数据清理
                ////baseList.DataBodyRange.ClearContents();
                ////基础List行数刷新
                ////int newRowCount = copyArray.GetLength(0);
                ////baseList.Resize(baseList.Range.Resize[newRowCount + 1, baseList.Range.Columns.Count]);
                ////baseList.DataBodyRange.Value2 = copyArray;

                ////基础标记数据删除
                ////var sheet = Wk.Worksheets["LTE【基础】"];
                ////var oldTagRange = sheet.Range["A2:A10000"];
                ////oldTagRange.Value2 = null;
                ////基础标记数据写入
                ////var tagRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[copyArray.GetLength(0) + 1, 1]];
                ////tagRange.Value2 = "+";

                //非Com写入数据,索引从0开始,效率确实更高,读取还是ListObject更方便
                var sheetName = "LTE【基础】";
                var rowMax = copyArray.GetLength(0);

                PubMetToExcel.WriteExcelDataC(
                    sheetName,
                    1,
                    ClearToLastRow,
                    BaseDataStartCol,
                    colCount,
                    null
                );
                PubMetToExcel.WriteExcelDataC(
                    sheetName,
                    1,
                    rowMax,
                    BaseDataStartCol,
                    colCount,
                    copyArray
                );

                PubMetToExcel.WriteExcelDataC(sheetName, 0, 0, BaseDataStartCol, colCount, null);
                PubMetToExcel.WriteExcelDataC(
                    sheetName,
                    0,
                    0,
                    BaseDataStartCol,
                    colCount,
                    copyTilteArray
                );

                baseList.Resize(baseList.Range.Resize[rowMax + 1, baseList.Range.Columns.Count]);

                PubMetToExcel.WriteExcelDataC(
                    sheetName,
                    1,
                    ClearToLastRow,
                    BaseDataTagCol,
                    BaseDataTagCol,
                    null
                );
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

                var fieldGroupList = RequireListObject("#道具信息", "道具信息");
                if (fieldGroupList == null)
                    return;
                object[,] fieldGroupArray = GetBodyRange(fieldGroupList, "道具信息");
                if (fieldGroupArray == null)
                    return;

                // 寻找优先级数据
                var findRanklistObjectsDic = PubMetToExcel.GetExcelListObjects(
                    WkPath,
                    "#【A-LTE】数值大纲.xlsx",
                    "#寻找优先级"
                );
                var findRankDataArray = GetTableData(
                    findRanklistObjectsDic,
                    "#寻找优先级",
                    "寻找方案"
                );
                if (findRankDataArray == null)
                    return;

                var baseActivityIdArray = PubMetToExcel.ReadExcelDataC("LTE【基础】", 0, 0, 1, 1);
                double baseActivityId = baseActivityIdArray?[0, 0] is double d ? d : 0;

                //寻找数据整理
                var findArray = FindData(
                    copyArray,
                    dataTypeArray,
                    fieldGroupArray,
                    copyTilteArray,
                    findRankDataArray,
                    baseActivityId
                );
                var sheetFindName = "LTE【寻找】";
                var rowFindMax = findArray.GetLength(0);

                PubMetToExcel.WriteExcelDataC(
                    sheetFindName,
                    1,
                    ClearToLastRow,
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

                findList.Resize(
                    findList.Range.Resize[rowFindMax + 1, findList.Range.Columns.Count]
                );

                object[,] writeFindArray = new object[rowFindMax, 1];
                for (int i = 0; i < rowFindMax; i++)
                    writeFindArray[i, 0] = "+";
                PubMetToExcel.WriteExcelDataC(
                    sheetFindName,
                    1,
                    ClearToLastRow,
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
        );

    //更新写入数据（指定范围内数据去重），比对数据，更新数据状态
    public static void UpdateCopyValue(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(UpdateCopyValue),
            () =>
            {
                object[,] copyArray = FilterRepeatValue(ActivityDataMinIndex, ActivityDataMaxIndex);
                object[,] copyTitleArray = ColTitleValue(
                    ActivityDataMinIndex,
                    ActivityDataMaxIndex
                );

                var list = RequireListObject("LTE【基础】", "LTE【基础】");
                if (list == null)
                    return;
                var findList = RequireListObject("LTE【寻找】", "LTE【寻找】");
                if (findList == null)
                    return;

                //基础数据修改依赖数据
                var listObjectsDic = PubMetToExcel.GetExcelListObjects(
                    WkPath,
                    "#【A-LTE】数值大纲.xlsx",
                    "#各类枚举"
                );
                var dataTypeArray = GetTableData(listObjectsDic, "#各类枚举", "数据类型");
                if (dataTypeArray == null)
                    return;

                //基础数据整理
                var copyData = BaseData(copyArray, dataTypeArray, copyTitleArray);
                copyArray = copyData.fixArray;
                copyTitleArray = copyData.fixTitleArray;
                var colCount = copyTitleArray.Length;
                var errorTypeList = copyData.errorTypeList;

                if (errorTypeList.Count != 0)
                {
                    var emptyKeyErrors = errorTypeList.Where(e => e.StartsWith('[')).ToList();
                    var unknownTypeErrors = errorTypeList.Where(e => !e.StartsWith('[')).ToList();
                    var sb = new StringBuilder("基础数据存在以下问题，请修复后重试：\n");
                    if (emptyKeyErrors.Count > 0)
                        sb.AppendLine(
                            $"\n【类型字段为空】共 {emptyKeyErrors.Count} 行：\n"
                                + string.Join("\n", emptyKeyErrors)
                        );
                    if (unknownTypeErrors.Count > 0)
                    {
                        var unknownOnly = new HashSet<string>(unknownTypeErrors);
                        sb.AppendLine($"\n【类型不在枚举中】{string.Join(", ", unknownOnly)}");
                    }
                    MessageBox.Show(sb.ToString());
                    return;
                }

                WriteDymaicData(copyArray, list, "LTE【基础】", 1, colCount);

                PubMetToExcel.WriteExcelDataC(
                    "LTE【基础】",
                    0,
                    0,
                    BaseDataStartCol,
                    colCount,
                    null
                );
                PubMetToExcel.WriteExcelDataC(
                    "LTE【基础】",
                    0,
                    0,
                    BaseDataStartCol,
                    colCount,
                    copyTitleArray
                );

                var fieldGroupList = RequireListObject("#道具信息", "道具信息");
                if (fieldGroupList == null)
                    return;
                object[,] fieldGroupArray = GetBodyRange(fieldGroupList, "道具信息");
                if (fieldGroupArray == null)
                    return;

                // 寻找优先级数据
                var findRanklistObjectsDic = PubMetToExcel.GetExcelListObjects(
                    WkPath,
                    "#【A-LTE】数值大纲.xlsx",
                    "#寻找优先级"
                );
                var findRankDataArray = GetTableData(
                    findRanklistObjectsDic,
                    "#寻找优先级",
                    "寻找方案"
                );
                if (findRankDataArray == null)
                    return;

                //寻找数据整理
                var baseActivityIdArray2 = PubMetToExcel.ReadExcelDataC("LTE【基础】", 0, 0, 1, 1);
                double baseActivityId2 = baseActivityIdArray2?[0, 0] is double d2 ? d2 : 0;
                var findArray = FindData(
                    copyArray,
                    dataTypeArray,
                    fieldGroupArray,
                    copyTitleArray,
                    findRankDataArray,
                    baseActivityId2
                );
                WriteDymaicData(findArray, findList, "LTE【寻找】", 1, 9);
            }
        );

    // "图1-1" → "图1"；null 或无 "-" 时安全返回空字符串
    private static string MapPrefix(string val) =>
        string.IsNullOrEmpty(val) ? string.Empty : val.Split('-')[0];

    // baseDic[id][col]；id 不存在或列不足时返回空字符串而非崩溃
    private static string SafeGet(Dictionary<string, List<string>> dic, string id, int col)
    {
        if (id == null || !dic.TryGetValue(id, out var row) || col >= row.Count)
            return string.Empty;
        return row[col] ?? string.Empty;
    }

    // 从 GetExcelListObjects 字典里取指定 key；为 null 或 key 不存在时弹提示并返回 null
    private static object[,] GetTableData(
        Dictionary<string, object[,]> dic,
        string dicName,
        string key
    )
    {
        if (dic == null)
        {
            MessageBox.Show($"找不到数据表「{dicName}」，请确认文件已打开");
            return null;
        }
        if (!dic.TryGetValue(key, out var arr) || arr == null)
        {
            MessageBox.Show($"「{dicName}」中找不到名称为「{key}」的列表，请确认表名正确");
            return null;
        }
        return arr;
    }

    // 读 ListObject 数据区；DataBodyRange 为空时弹提示并返回 null
    private static object[,] GetBodyRange(ListObject list, string listDesc)
    {
        var range = list.DataBodyRange;
        if (range == null)
        {
            MessageBox.Show($"「{listDesc}」数据区为空，请先确认表中有数据行");
            return null;
        }
        if (range.Value2 is not object[,] val)
        {
            MessageBox.Show($"「{listDesc}」数据区读取失败（Value2 为空）");
            return null;
        }
        return val;
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
        object[,] oldListData = GetBodyRange(list, sheetName) ?? new object[0, 0];
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
        var excel = AppServices.App;

        var sheet = excel.ActiveSheet as Worksheet;

        var usedRange = sheet?.UsedRange;
        Debug.Assert(usedRange != null, nameof(usedRange) + " != null");
        // ReSharper disable once PossibleNullReferenceException
        var usedMaxRow = usedRange.Rows.Count;

        Range copyRange;
        if (!isSelect)
        {
            /*
                        if (false)
                        {
                            MessageBox.Show("未找到表");
                            return null;
                        }
            */
            // ReSharper disable once PossibleNullReferenceException
            var copyColMin = sheet.Range[min].Value2;
            var copyColMax = sheet.Range[max].Value2;
            copyRange = sheet.Range[
                sheet.Cells[3, copyColMin],
                sheet.Cells[usedMaxRow + 10, copyColMax]
            ];
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

    //指定列范围的数据[字段名]
    private static object[,] ColTitleValue(string min, string max)
    {
        var excel = AppServices.App;

        var sheet = excel.ActiveSheet as Worksheet;

        if (sheet is not null)
        {
            var copyColMin = sheet.Range[min].Value2;
            var copyColMax = sheet.Range[max].Value2;

            Range colTitleRange = sheet.Range[
                sheet.Cells[2, copyColMin],
                sheet.Cells[2, copyColMax]
            ];
            object[,] colTitleArray = colTitleRange.Value2;

            return colTitleArray;
        }
        return null;
    }

    //原始数据改造
    private static (
        object[,] fixArray,
        List<string> errorTypeList,
        object[,] fixTitleArray
    ) BaseData(object[,] baseArray, object[,] dataTypeArray, object[,] baseTilteArray)
    {
        var baseDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(baseArray);
        var dataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(dataTypeArray);
        var baseTitleDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(baseTilteArray);

        baseTitleDic["数据编号"].Add("寻找类型");
        baseTitleDic["数据编号"].Add("寻找细类");
        baseTitleDic["数据编号"].Add("链长");
        baseTitleDic["数据编号"].Add("五合提示");
        baseTitleDic["数据编号"].Add("消耗ID组");
        baseTitleDic["数据编号"].Add("产出ID组");
        baseTitleDic["数据编号"].Add("消耗量组");
        baseTitleDic["数据编号"].Add("产出量组");
        baseTitleDic["数据编号"].Add("转化ID组");

        var errorTypeList = new List<string>();

        foreach (var baseList in baseDic)
        {
            var key = baseList.Key;
            if (string.IsNullOrEmpty(key))
                continue;

            PluginLog.Write($"[LteData][BaseData] 数据编号：{key}");

            // 资源编号和图片编号
            string prefabId = baseDic[key][BaseDicColPrefabId];
            if (string.IsNullOrEmpty(prefabId))
            {
                baseDic[key][BaseDicColPrefabId] = key;
            }
            string iconId = baseDic[key][BaseDicColIconId];
            if (string.IsNullOrEmpty(iconId))
            {
                baseDic[key][BaseDicColIconId] = key;
            }

            //寻找类型、寻找细类
            string findType;
            string findDetailType;

            string itemType = baseDic[key][BaseDicColType];

            if (string.IsNullOrEmpty(itemType))
            {
                errorTypeList.Add($"[{key}] 类型字段为空");
                continue;
            }

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
            string currentName = SafeGet(baseDic, key, BaseDicColPackage);
            int countCurrent = baseDic
                .Values.Where(list => list.Count > BaseDicColPackage)
                .Count(list => list[BaseDicColPackage] == currentName);

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
            string rank = SafeGet(baseDic, key, BaseDicColLevel);

            if (int.TryParse(rank, out int rankNum))
            {
                if (rankNum >= 6 && rankNum < countCurrent)
                {
                    fiveMergeTip = "35";
                }
            }
            baseDic[key].Add(fiveMergeTip);

            //消耗ID组、产出ID组、消耗量组、产出量组

            var consumeIdList = new List<string>();
            var productIdList = new List<string>();
            var consumeCountList = new List<string>();
            var productCountList = new List<string>();

            var idNameList = new List<int> { 11, 13, 15, 17, 19, 21, 23, 25 };
            var countNumList = new List<int> { 12, 14, 16, 18, 20, 22, 24, 26 };

            string firstPos = SafeGet(baseDic, key, BaseDicColFirstMap);
            var firstPosPre = MapPrefix(firstPos);

            int onlyNum = BaseDicColOnlyName;
            int num = BaseDicColName;

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
                            //先在唯一ID中查找 name
                            matchId = baseDic
                                .FirstOrDefault(kv =>
                                    kv.Value.Count > onlyNum && kv.Value[onlyNum] == name
                                )
                                .Key;
                        }
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

            //如果没有消耗ID组则尝试查找谁消耗该ID(只针对确定需要寻找的物品）
            //用来建立道具关系,此时组成的消耗组，只是为了建立寻找关系

            if (consumeIdList.Count == 0)
            {
                // 该ID的类型是否需要寻找（针对修复物）
                var isFind = dataTypeDic[itemType][4];
                if (isFind == "1" && itemType.Contains("修"))
                {
                    // 该ID的代号和唯一代号
                    var orginNum = baseDic[key][num];
                    var orginOnlyNum = baseDic[key][onlyNum];
                    string matchId = baseDic
                        .FirstOrDefault(kv =>
                            kv.Value.Count > 17
                            && MapPrefix(kv.Value[BaseDicColFirstMap]) + kv.Value[11]
                                == orginOnlyNum
                        )
                        .Key;
                    if (matchId == null)
                    {
                        matchId = baseDic
                            .FirstOrDefault(kv =>
                                kv.Value.Count > 17
                                && MapPrefix(kv.Value[BaseDicColFirstMap]) + kv.Value[13]
                                    == orginOnlyNum
                            )
                            .Key;
                    }
                    if (matchId == null)
                    {
                        matchId = baseDic
                            .FirstOrDefault(kv =>
                                kv.Value.Count > 17
                                && MapPrefix(kv.Value[BaseDicColFirstMap]) + kv.Value[15]
                                    == orginOnlyNum
                            )
                            .Key;
                    }
                    if (matchId == null)
                    {
                        matchId = baseDic
                            .FirstOrDefault(kv =>
                                kv.Value.Count > 17
                                && MapPrefix(kv.Value[BaseDicColFirstMap]) + kv.Value[17]
                                    == orginOnlyNum
                            )
                            .Key;
                    }
                    //唯一代号找不到尝试寻找代号
                    if (matchId == null)
                    {
                        matchId = baseDic
                            .FirstOrDefault(kv => kv.Value.Count > 11 && kv.Value[11] == orginNum)
                            .Key;
                    }
                    if (matchId != null)
                    {
                        consumeIdList.Add(matchId);
                        consumeCountList.Add("1");
                    }
                }
            }

            var consumeIdGroup = string.Join("#", consumeIdList);
            var productIdGroup = string.Join("#", productIdList);
            var consumeCountGroup = string.Join("#", consumeCountList);
            var productCountGroup = string.Join("#", productCountList);

            baseDic[key].Add(consumeIdGroup);
            baseDic[key].Add(productIdGroup);
            baseDic[key].Add(consumeCountGroup);
            baseDic[key].Add(productCountGroup);

            // 转换组
            var spawnIndex = 27;
            var spawnName = SafeGet(baseDic, key, spawnIndex);

            //先在唯一ID中查找
            var spawnMatchId = string.Empty;

            if (!string.IsNullOrEmpty(spawnName))
            {
                spawnMatchId = baseDic
                    .FirstOrDefault(kv =>
                        kv.Value.Count > onlyNum && kv.Value[onlyNum] == firstPosPre + spawnName
                    )
                    .Key;
                if (spawnMatchId == null)
                {
                    //先在唯一ID中查找 name
                    spawnMatchId = baseDic
                        .FirstOrDefault(kv =>
                            kv.Value.Count > onlyNum && kv.Value[onlyNum] == spawnName
                        )
                        .Key;
                }
                if (spawnMatchId == null)
                {
                    //后在ID中查找
                    spawnMatchId = baseDic
                        .FirstOrDefault(kv => kv.Value.Count > num && kv.Value[num] == spawnName)
                        .Key;
                }
            }

            baseDic[key].Add(spawnMatchId);
        }
        var fixArray = PubMetToExcel.DictionaryTo2DArray(
            baseDic,
            baseDic.Count,
            baseDic[baseDic.Keys.First()].Count
        );

        var fixTitleArray = PubMetToExcel.DictionaryTo2DArray(
            baseTitleDic,
            baseTitleDic.Count,
            baseTitleDic[baseTitleDic.Keys.First()].Count
        );
        return (fixArray, errorTypeList, fixTitleArray);
    }
    #endregion

    #region LTE寻找数据计算
    private static object[,] FindData(
        object[,] copyArray,
        object[,] dataTypeArray,
        object[,] fieldGroupArray,
        object[,] copyTitleArray,
        object[,] findRankDataArray,
        double activityId = 0
    )
    {
        var findDic = new Dictionary<string, List<string>>();

        var copyDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(copyArray);
        var dataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(dataTypeArray);
        var fieldGroupDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(fieldGroupArray);
        var copyTitleDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(copyTitleArray);

        var findRankDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(findRankDataArray);

        var titleList = copyTitleDic["数据编号"];

        foreach (var key in copyDic.Keys)
        {
            var keyType = copyDic[key][titleList.IndexOf("类型")];
            if (!dataTypeDic.ContainsKey(keyType))
            {
                continue;
            }
            var dataType = dataTypeDic[keyType][4];
            if (dataType != "1")
            {
                continue;
            }

            // 消耗组，第1层寻找
            var inputGroup = copyDic[key][titleList.IndexOf("消耗ID组")];
            if (string.IsNullOrEmpty(inputGroup))
                continue;
            var inputArray = inputGroup.Split('#');

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

                if (!copyDic.ContainsKey(findTargetId))
                    continue;
                var findTargetType = copyDic[findTargetId][titleList.IndexOf("寻找类型")];
                var findTargetDetailType = copyDic[findTargetId][titleList.IndexOf("寻找细类")];

                if (findTargetType != String.Empty)
                {
                    var findLinksGroup = FindLinks(
                        findTargetDetailType,
                        findTargetType,
                        findTargetId,
                        copyDic,
                        out var findTips,
                        fieldGroupDic,
                        titleList,
                        findRankDic,
                        dataTypeDic
                    );
                    var findLinks = findLinksGroup.findLinks;
                    var findLinks31 = findLinksGroup.findLinks31;

                    if (findLinks != String.Empty)
                    {
                        var findIdStr = Convert.ToString(findId, CultureInfo.InvariantCulture);
                        var onlyName = copyDic[findIdStr][titleList.IndexOf("唯一代号")];
                        if (onlyName != String.Empty)
                        {
                            if (!findDic.ContainsKey(findIdStr))
                            {
                                findDic.Add(findIdStr, new List<string>());
                                findDic[findIdStr].Add(findIdStr);
                                findDic[findIdStr]
                                    .Add(copyDic[findIdStr][titleList.IndexOf("首次出现")]);
                                findDic[findIdStr]
                                    .Add(copyDic[findIdStr][titleList.IndexOf("唯一代号")]);
                                findDic[findIdStr]
                                    .Add(copyDic[findIdStr][titleList.IndexOf("代号")]);
                                findDic[findIdStr]
                                    .Add(copyDic[findIdStr][titleList.IndexOf("当前包装")]);
                                findDic[findIdStr]
                                    .Add("寻-" + copyDic[key][titleList.IndexOf("类型")]);
                                findDic[findIdStr]
                                    .Add(copyDic[findIdStr][titleList.IndexOf("备注名称")]);
                                findDic[findIdStr].Add(findTips);

                                // 追加地图节点：从目标物品"首次出现"取地图ID，与任务寻找逻辑一致
                                var mapName = copyDic[findIdStr][titleList.IndexOf("首次出现")];
                                var mapPrefix = MapPrefix(mapName);
                                var mapMatch = _digitsRegex.Match(mapPrefix);
                                var mapId = "0";
                                if (
                                    mapMatch.Success
                                    && double.TryParse(mapMatch.Value, out double mapIdDouble)
                                )
                                    mapId = (mapIdDouble + activityId).ToString(
                                        CultureInfo.InvariantCulture
                                    );

                                var mapEntrance =
                                    mapId != "0"
                                        ? FindLinkMapEntrance + mapId + FindLinkEndPortal
                                        : string.Empty;
                                findDic[findIdStr].Add(findLinks + findLinks31 + mapEntrance);
                            }
                        }
                    }
                }
                //}
            }

            ////反向查找
            //List<string> subMatchIDs = copyDic
            //    .Where(kv => kv.Value.Count > 30 && kv.Value[30].Contains(key))
            //    .Select(kv => kv.Key)
            //    .ToList();

            //var subFindLinks = string.Empty;
            //foreach (var findTargetId2 in subMatchIDs)
            //{
            //    if (findTargetId2 != string.Empty)
            //    {
            //        var subFindId = findTargetId2;
            //        var findTargetType2 = copyDic[findTargetId2][25];
            //        var findTargetDetailType2 = copyDic[findTargetId2][26];

            //        if (findTargetType2 != string.Empty)
            //        {
            //            if (findTargetDetailType2 == string.Empty)
            //            {
            //                findTargetDetailType2 = "未找到细类";
            //            }
            //            if (findTargetType2 == "19")
            //            {
            //                subFindLinks +=
            //                    "{"
            //                    + findTargetType2
            //                    + ","
            //                    + findTargetDetailType2
            //                    + ","
            //                    + findTargetId2
            //                    + "},";
            //            }
            //            else if (findTargetType2 == "1")
            //            {
            //                subFindLinks += "{" + findTargetType2 + "," + findTargetId2 + "},";
            //            }
            //            else
            //            {
            //                subFindLinks += "{" + findTargetType2 + "," + findTargetId2 + "},";
            //            }
            //        }
            //        if (subFindLinks != String.Empty)
            //        {
            //            if (!findDic.ContainsKey(subFindId))
            //            {
            //                findDic.Add(subFindId, new List<string>());
            //                findDic[subFindId].Add(subFindId);
            //                findDic[subFindId].Add(copyDic[subFindId][1]);
            //                findDic[subFindId].Add(copyDic[subFindId][2]);
            //                findDic[subFindId].Add(copyDic[subFindId][3]);
            //                findDic[subFindId].Add(copyDic[subFindId][4]);
            //                findDic[subFindId].Add("寻-" + copyDic[key][6]);
            //                findDic[subFindId].Add(copyDic[subFindId][7]);
            //                findDic[subFindId].Add("");

            //                var findLinksFix = subFindLinks.Substring(0, subFindLinks.Length - 1);
            //                findLinksFix += ",{8,9999}";

            //                findDic[subFindId].Add(findLinksFix);
            //            }
            //        }
            //    }
            //}
        }
        var findLinksArray = PubMetToExcel.DictionaryTo2DArray(
            findDic,
            findDic.Count,
            findDic[findDic.Keys.First()].Count
        );
        return findLinksArray;
    }

    private static (string findLinks, string findLinks31) FindLinks(
        string findTargetDetailType,
        string findTargetType,
        string findTargetId,
        Dictionary<string, List<string>> baseDic,
        out string findTips,
        Dictionary<string, List<string>> fieldGroupDic,
        List<string> titleList,
        Dictionary<string, List<string>> findRankDic,
        Dictionary<string, List<string>> dataTypeDic
    )
    {
        var findLinks = string.Empty;
        var findLinks31 = string.Empty;

        findTips = string.Empty;

        var findTargetDetailTypeIndex = titleList.IndexOf("寻找细类");
        var findTargetTypeIndex = titleList.IndexOf("寻找类型");

        if (!baseDic.ContainsKey(findTargetId))
        {
            findTips = string.Empty;
            return (string.Empty, string.Empty);
        }

        var findTargetNickName = SafeGet(baseDic, findTargetId, titleList.IndexOf("代号"));
        var findTargetOnlyNickName = SafeGet(baseDic, findTargetId, titleList.IndexOf("唯一代号"));
        var findTaregtfieldLinks = FieldGroupLinks(
            fieldGroupDic,
            findTargetNickName,
            findTargetOnlyNickName
        );
        var targetTypeIndex = titleList.IndexOf("类型");
        var targetType = SafeGet(baseDic, findTargetId, targetTypeIndex);

        if (findRankDic.TryGetValue(targetType, out var findTargetRankList))
        {
            (findLinks, findLinks31, findTips) = FindLinksWithRank(
                findTargetDetailType,
                findTargetId,
                baseDic,
                titleList,
                dataTypeDic,
                findRankDic,
                findTargetRankList,
                fieldGroupDic,
                targetTypeIndex,
                findTargetDetailTypeIndex,
                findTargetTypeIndex
            );
        }
        else
        {
            (findLinks, findLinks31, findTips) = FindLinksWithoutRank(
                findTargetDetailType,
                findTargetType,
                findTargetId,
                targetType,
                baseDic,
                titleList,
                fieldGroupDic,
                findTaregtfieldLinks
            );
        }
        // 去重
        findLinks = RemoveDuplicateBracketsLinqOrdered(findLinks);

        findLinks31 = RemoveDuplicateBracketsLinqOrdered(findLinks31);

        if (findLinks != string.Empty)
        {
            findLinks += ",";
        }
        if (findLinks31 != string.Empty)
        {
            findLinks31 += ",";
        }

        return (findLinks, findLinks31);
    }

    // 有优先级寻找：按 findRankDic 定义的层级递归展开，最多 2 层
    private static (string findLinks, string findLinks31, string findTips) FindLinksWithRank(
        string findTargetDetailType,
        string findTargetId,
        Dictionary<string, List<string>> baseDic,
        List<string> titleList,
        Dictionary<string, List<string>> dataTypeDic,
        Dictionary<string, List<string>> findRankDic,
        List<string> findTargetRankList,
        Dictionary<string, List<string>> fieldGroupDic,
        int targetTypeIndex,
        int findTargetDetailTypeIndex,
        int findTargetTypeIndex
    )
    {
        var findLinks = string.Empty;
        var findLinks31 = string.Empty;
        var findTips = string.Empty;

        // 1层寻找
        var findRankLinks1 = FindRankLinks(
            findTargetId,
            baseDic,
            titleList,
            dataTypeDic,
            findTargetRankList,
            targetTypeIndex,
            findTargetDetailTypeIndex,
            findTargetTypeIndex
        );

        var findItemGroup = findRankLinks1
            .Where(x => x.targetFindDetailType != "4")
            .Select(x => x.targetId)
            .ToList();

        if (findRankLinks1.Count > 0)
        {
            findLinks += LinksBuild(findRankLinks1, fieldGroupDic, baseDic, titleList).findLinks;
            findLinks31 += LinksBuild(findRankLinks1, null, null, null).findLinks31;
        }

        // 2层寻找
        var findRankLinks2 = new List<(string, string, string, string)>();
        foreach (var rankLinks1 in findRankLinks1)
        {
            if (rankLinks1.hasFind != "1")
            {
                var target2Id = rankLinks1.targetId;
                if (target2Id == findTargetId)
                    continue;
                if (!baseDic.ContainsKey(target2Id))
                    continue;
                var target2Type = SafeGet(baseDic, target2Id, targetTypeIndex);
                if (findRankDic.TryGetValue(target2Type, out var findTarget2RankList))
                {
                    var findRankLinksTemp = FindRankLinks(
                        target2Id,
                        baseDic,
                        titleList,
                        dataTypeDic,
                        findTarget2RankList,
                        targetTypeIndex,
                        findTargetDetailTypeIndex,
                        findTargetTypeIndex
                    );
                    var findItemGroupTemp = findRankLinksTemp
                        .Where(x => x.targetFindDetailType != "4")
                        .Select(x => x.targetId)
                        .ToList();
                    findRankLinks2.AddRange(findRankLinksTemp);
                    findItemGroup.AddRange(findItemGroupTemp);
                }
            }
        }

        if (findRankLinks2.Count > 0)
        {
            findLinks += LinksBuild(findRankLinks2, fieldGroupDic, baseDic, titleList).findLinks;
            findLinks31 += LinksBuild(findRankLinks2, null, null, null).findLinks31;
        }

        // 寻找界面提示使用最后的id，因为其他id可能没有图片资源
        var finalMatchId = findItemGroup.FirstOrDefault();
        if (findRankLinks1.Count == 0 || finalMatchId is null or "")
        {
            findTips = "{1,\"tip_obstacleItem\",2}";
        }
        else
        {
            if (findTargetDetailType == "4")
            {
                findTips =
                    "{3,"
                    + findTargetId.Substring(0, findTargetId.Length - 2)
                    + "00,"
                    + finalMatchId
                    + "}";
            }
            else
            {
                findTips = "{1,\"tip_obstacleItem\",1," + finalMatchId + "}";
            }
        }

        return (findLinks, findLinks31, findTips);
    }

    // 传统寻找（无优先级）：按产出ID组向上逐层查找，最多 3 层
    private static (string findLinks, string findLinks31, string findTips) FindLinksWithoutRank(
        string findTargetDetailType,
        string findTargetType,
        string findTargetId,
        string targetType,
        Dictionary<string, List<string>> baseDic,
        List<string> titleList,
        Dictionary<string, List<string>> fieldGroupDic,
        string findTaregtfieldLinks
    )
    {
        var findLinks = string.Empty;
        var findLinks31 = string.Empty;
        var findTips = string.Empty;

        PluginLog.Write($"[LteData][FindLinks] 【#寻找优先级】表中不存在：{targetType}");

        //1层查找
        if (findTargetDetailType == string.Empty)
            findTargetDetailType = "未找到细类";

        if (findTargetType == "19")
        {
            findLinks +=
                "{" + findTargetType + "," + findTargetDetailType + "," + findTargetId + "},";
        }
        else if (findTargetType == "1")
        {
            findLinks += "{" + findTargetType + "," + findTargetId + "},";
            findLinks31 += "{" + 31 + "," + findTargetId + "},";
        }
        else
        {
            findLinks += "{" + findTargetType + "," + findTargetId + "},";
        }

        if (findTaregtfieldLinks != string.Empty)
            findLinks += findTaregtfieldLinks;

        int outPutIdGroupIndex = titleList.IndexOf("产出ID组");
        // 2层查找
        List<string> matchedIDsOri = baseDic
            .Where(kv =>
                kv.Value.Count > outPutIdGroupIndex
                && kv.Value[outPutIdGroupIndex].Contains(findTargetId)
            )
            .Select(kv => kv.Key)
            .ToList();

        // 没有直接匹配的，需要继续查找（按照链的规则，要验证类型）
        List<string> matchedIDsEnd = new();
        if (targetType.Contains("链"))
        {
            var findTargetId01 = findTargetId.Substring(0, findTargetId.Length - 2) + "01";
            var findTargetId02 = findTargetId.Substring(0, findTargetId.Length - 2) + "02";
            var findTargetId03 = findTargetId.Substring(0, findTargetId.Length - 2) + "03";
            foreach (
                var variantId in new[]
                {
                    findTargetId,
                    findTargetId01,
                    findTargetId02,
                    findTargetId03,
                }
            )
            {
                matchedIDsEnd.AddRange(
                    baseDic
                        .Where(kv =>
                            kv.Value.Count > outPutIdGroupIndex
                            && kv.Value[outPutIdGroupIndex].Contains(variantId)
                        )
                        .Select(kv => kv.Key)
                );
            }
        }

        // 3层查找
        if (matchedIDsEnd.Count > 0)
        {
            var layer3 = new List<string>();
            foreach (var findTargetId2 in matchedIDsEnd)
            {
                layer3.AddRange(
                    baseDic
                        .Where(kv =>
                            kv.Value.Count > outPutIdGroupIndex
                            && kv.Value[outPutIdGroupIndex].Contains(findTargetId2)
                        )
                        .Select(kv => kv.Key)
                );
            }
            matchedIDsEnd.AddRange(layer3);
        }

        // 寻找字符串格式化（去重）
        List<string> matchedIDs = new HashSet<string>(matchedIDsEnd).ToList();
        string finalMatchedId = matchedIDs.LastOrDefault() ?? string.Empty;

        if (matchedIDs.Count == 0)
        {
            findTips = "{1,\"tip_obstacleItem\",2}";
        }
        else
        {
            int itemCount = 0;
            foreach (var findTargetId3 in matchedIDs)
            {
                if (findTargetId3 != string.Empty && baseDic.ContainsKey(findTargetId3))
                {
                    var findTargetType3 = SafeGet(
                        baseDic,
                        findTargetId3,
                        titleList.IndexOf("寻找类型")
                    );
                    var findTargetDetailType3 = SafeGet(
                        baseDic,
                        findTargetId3,
                        titleList.IndexOf("寻找细类")
                    );
                    var findTargetNickName3 = SafeGet(
                        baseDic,
                        findTargetId3,
                        titleList.IndexOf("代号")
                    );
                    var findTargetOnlyNickName3 = SafeGet(
                        baseDic,
                        findTargetId,
                        titleList.IndexOf("唯一代号")
                    );
                    var findTaregtfieldLinks3 = FieldGroupLinks(
                        fieldGroupDic,
                        findTargetNickName3,
                        findTargetOnlyNickName3
                    );

                    if (findTargetType3 != string.Empty)
                    {
                        if (findTargetDetailType3 == string.Empty)
                            findTargetDetailType3 = "未找到细类";

                        if (findTargetType3 == "19")
                        {
                            findLinks +=
                                "{"
                                + findTargetType3
                                + ","
                                + findTargetDetailType3
                                + ","
                                + findTargetId3
                                + "},";
                        }
                        else if (findTargetType3 == "1")
                        {
                            findLinks += "{" + findTargetType3 + "," + findTargetId3 + "},";
                            findLinks31 += "{" + 31 + "," + findTargetId3 + "},";
                        }
                        else
                        {
                            findLinks += "{" + findTargetType3 + "," + findTargetId3 + "},";
                        }

                        if (findTaregtfieldLinks3 != string.Empty)
                            findLinks += findTaregtfieldLinks3;

                        if (itemCount == 0)
                        {
                            findTips =
                                findTargetDetailType == "4"
                                    ? "{3,"
                                        + findTargetId.Substring(0, findTargetId.Length - 2)
                                        + "00,"
                                        + finalMatchedId
                                        + "}"
                                    : "{1,\"tip_obstacleItem\",1," + finalMatchedId + "}";
                        }
                    }
                }
                itemCount++;
            }
        }

        return (findLinks, findLinks31, findTips);
    }

    private static (string findLinks, string findLinks31) LinksBuild(
        List<(
            string targetId,
            string targetFindType,
            string targetFindDetailType,
            string hasFind
        )> findRankLinks,
        Dictionary<string, List<string>> fieldGroupDic,
        Dictionary<string, List<string>> baseDic,
        List<string> titleList
    )
    {
        var findLinks = string.Empty;
        var findLinks31 = string.Empty;

        foreach (var rankLinks in findRankLinks)
        {
            if (rankLinks.targetFindType == "1")
            {
                findLinks += "{" + rankLinks.targetFindType + "," + rankLinks.targetId + "},";
                findLinks31 += "{31," + rankLinks.targetId + "},";
            }
            else if (rankLinks.targetFindType == "19")
            {
                findLinks +=
                    "{"
                    + rankLinks.targetFindType
                    + ","
                    + rankLinks.targetFindDetailType
                    + ","
                    + rankLinks.targetId
                    + "},";
            }
            else if (rankLinks.targetFindType == "18")
            {
                findLinks += "{" + rankLinks.targetFindType + "," + rankLinks.targetId + "},";
            }
            if (fieldGroupDic != null && baseDic.ContainsKey(rankLinks.targetId))
            {
                var findTargetNickName = SafeGet(
                    baseDic,
                    rankLinks.targetId,
                    titleList.IndexOf("代号")
                );
                var findTargetOnlyNickName = SafeGet(
                    baseDic,
                    rankLinks.targetId,
                    titleList.IndexOf("唯一代号")
                );
                var findTaregtfieldLinks = FieldGroupLinks(
                    fieldGroupDic,
                    findTargetNickName,
                    findTargetOnlyNickName
                );

                if (findTaregtfieldLinks != string.Empty)
                    findLinks += findTaregtfieldLinks;
            }
        }

        return (findLinks, findLinks31);
    }

    private static List<(
        string targetId,
        string targetFindType,
        string targetFindDetailType,
        string hasFind
    )> FindRankLinks(
        string findTargetId,
        Dictionary<string, List<string>> baseDic,
        List<string> titleList,
        Dictionary<string, List<string>> dataTypeDic,
        List<string> findTargetRankList,
        int targetTypeIndex,
        int findTargetDetailTypeIndex,
        int findTargetTypeIndex
    )
    {
        var findItemListTuple = new List<(string, string, string, string)>();

        if (!baseDic.ContainsKey(findTargetId))
            return findItemListTuple;
        var targetType = SafeGet(baseDic, findTargetId, targetTypeIndex);

        // 优先级寻找
        foreach (var findTargetRank in findTargetRankList.Skip(1))
        {
            if (findTargetRank == null)
                continue;

            var rankParts = findTargetRank.Split('#');

            var findRankType = rankParts[0];
            var findRankParams1 = string.Empty;
            if (rankParts.Length >= 2)
            {
                findRankParams1 = rankParts[1];
            }

            int outPutIdGroupIndex = titleList.IndexOf("产出ID组");

            List<string> targetIdGroup;

            if (findRankType != "Any")
            {
                targetIdGroup = baseDic
                    .Where(kv =>
                        kv.Value.Count > outPutIdGroupIndex
                        && kv.Value[outPutIdGroupIndex].Contains(findTargetId)
                        && kv.Value[targetTypeIndex] == findRankType
                    )
                    .Select(kv => kv.Key)
                    .ToList();
            }
            else
            {
                targetIdGroup = baseDic
                    .Where(kv =>
                        kv.Value.Count > outPutIdGroupIndex
                        && kv.Value[outPutIdGroupIndex].Contains(findTargetId)
                    )
                    .Select(kv => kv.Key)
                    .ToList();
            }
            // 来源：非计算类型（查询）
            if (targetIdGroup.Count > 0)
            {
                foreach (var targetIdSource in targetIdGroup)
                {
                    var targetId = targetIdSource;
                    if (findRankParams1 != string.Empty)
                    {
                        if (findRankType.Contains("地标"))
                        {
                            var targetLevel = baseDic[targetIdSource][
                                titleList.IndexOf(findRankParams1)
                            ];
                            targetId = (
                                Convert.ToDouble(targetId) - Convert.ToDouble(targetLevel) + 1
                            ).ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    var targetFindType = baseDic[targetId][findTargetTypeIndex];
                    var targetFindDetailType = baseDic[targetId][findTargetDetailTypeIndex];
                    var targetTypeNew = baseDic[targetId][targetTypeIndex];
                    if (!dataTypeDic.TryGetValue(targetTypeNew, out var targetTypeRow))
                    {
                        PluginLog.Write(
                            $"[LteData][FindItemListTuple] 类型 {targetTypeNew} 不在 dataTypeDic，跳过 id={targetId}"
                        );
                        continue;
                    }
                    var hasFind = targetTypeRow[4];

                    findItemListTuple.Add(
                        (targetId, targetFindType, targetFindDetailType, hasFind)
                    );
                }
            }
            // 来源：计算类型
            else
            {
                var targetId = string.Empty;

                // 来源：链-建类型需要计算ID
                if (findRankType == "链-建")
                {
                    targetId = (Convert.ToDouble(findTargetId) + 10).ToString(
                        CultureInfo.CurrentCulture
                    );
                    if (!baseDic.ContainsKey(targetId))
                        targetId = (Convert.ToDouble(findTargetId) + 70).ToString(
                            CultureInfo.CurrentCulture
                        );
                    if (!baseDic.ContainsKey(targetId))
                    {
                        targetId = string.Empty;
                    }
                    else
                    {
                        var targetTypeNew = baseDic[targetId][targetTypeIndex];
                        if (!targetTypeNew.Contains("链"))
                        {
                            targetId = string.Empty;
                        }
                    }
                }
                // 来源：链-网需要寻找上1个网Id或不变
                else if (findRankType == "链-网")
                {
                    targetId = (
                        Convert.ToDouble(findTargetId) + 30 + Convert.ToDouble(findRankParams1)
                    ).ToString(CultureInfo.CurrentCulture);

                    PluginLog.Write(
                        $"[LteData][FindRankLinks] 谁找1：{findTargetId},找谁：{targetId}"
                    );

                    if (!baseDic.ContainsKey(targetId))
                    {
                        targetId = (
                            Convert.ToDouble(findTargetId) + 60 + Convert.ToDouble(findRankParams1)
                        ).ToString(CultureInfo.CurrentCulture);

                        PluginLog.Write(
                            $"[LteData][FindRankLinks] 谁找2：{findTargetId},找谁：{targetId}"
                        );

                        if (!baseDic.ContainsKey(targetId))
                            targetId = (
                                Convert.ToDouble(findTargetId)
                                + 80
                                + Convert.ToDouble(findRankParams1)
                            ).ToString(CultureInfo.CurrentCulture);

                        PluginLog.Write(
                            $"[LteData][FindRankLinks] 谁找3：{findTargetId},找谁：{targetId}"
                        );

                        if (!baseDic.ContainsKey(targetId))
                        {
                            targetId = string.Empty;
                        }
                        else
                        {
                            var targetTypeNew = baseDic[targetId][targetTypeIndex];
                            if (!targetTypeNew.Contains("链"))
                            {
                                targetId = string.Empty;
                            }
                        }

                        PluginLog.Write(
                            $"[LteData][FindRankLinks] 谁找4：{findTargetId},找谁：{targetId}"
                        );
                    }
                }
                // 来源：Any-链-源 — 追溯链的最低级版本(01)的产出来源
                // 若当前链后两位已是01，跳过避免与 Any/类型匹配规则重复
                else if (findRankType == "Any-链-源")
                {
                    var id01 = findTargetId.Length >= 2 ? findTargetId[..^2] + "01" : string.Empty;

                    if (!string.IsNullOrEmpty(id01) && id01 != findTargetId)
                    {
                        var sourceIds = baseDic
                            .Where(kv =>
                                kv.Value.Count > outPutIdGroupIndex
                                && kv.Value[outPutIdGroupIndex].Contains(id01)
                            )
                            .Select(kv => kv.Key)
                            .ToList();

                        foreach (var sourceId in sourceIds)
                        {
                            var sourceFindType = SafeGet(baseDic, sourceId, findTargetTypeIndex);
                            var sourceFindDetailType = SafeGet(
                                baseDic,
                                sourceId,
                                findTargetDetailTypeIndex
                            );
                            var sourceType = SafeGet(baseDic, sourceId, targetTypeIndex);
                            if (!dataTypeDic.ContainsKey(sourceType))
                                continue;
                            var hasFind = dataTypeDic[sourceType][4];
                            findItemListTuple.Add(
                                (sourceId, sourceFindType, sourceFindDetailType, hasFind)
                            );
                        }
                    }
                    continue;
                }
                else if (findRankType == targetType)
                {
                    targetId = findTargetId;
                }

                if (targetId != string.Empty)
                {
                    var targetFindType = baseDic[targetId][findTargetTypeIndex];
                    var targetFindDetailType = baseDic[targetId][findTargetDetailTypeIndex];
                    var targetTypeNew = baseDic[targetId][targetTypeIndex];
                    if (!dataTypeDic.TryGetValue(targetTypeNew, out var targetTypeRow))
                    {
                        PluginLog.Write(
                            $"[LteData][FindItemListTuple] 类型 {targetTypeNew} 不在 dataTypeDic，跳过 id={targetId}"
                        );
                        continue;
                    }
                    var hasFind = targetTypeRow[4];

                    findItemListTuple.Add(
                        (targetId, targetFindType, targetFindDetailType, hasFind)
                    );
                }
            }
        }

        return findItemListTuple;
    }

    public static string RemoveDuplicateBracketsLinqOrdered(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        var matches = Regex.Matches(input, @"{[^}]+}");

        var uniqueItems = matches.Select(m => m.Value).Distinct().ToList();

        var returnValue = string.Join(",", uniqueItems);
        return returnValue;
    }

    private static string FieldGroupLinks(
        Dictionary<string, List<string>> fieldGroupDic,
        string findTargetNickName,
        string findTargetOnlyNickName
    )
    {
        string fieldLinks = String.Empty;
        var fieldList = new List<string>();

        if (fieldGroupDic.ContainsKey(findTargetNickName))
        {
            fieldList = fieldGroupDic[findTargetNickName];
        }
        else
        {
            if (fieldGroupDic.ContainsKey(findTargetOnlyNickName))
            {
                fieldList = fieldGroupDic[findTargetOnlyNickName];
            }
        }
        if (fieldList.Count > 2)
        {
            var fieldMap = fieldList[1];
            var fieldPrefix = OutputWildcardPubDic["活动地组"];
            double fieldIndex = 0;
            if (double.TryParse(fieldMap, out double fileMapDouble))
            {
                if (double.TryParse(fieldPrefix, out double fieldPrefixDouble))
                {
                    fieldIndex = fileMapDouble * 100 + fieldPrefixDouble;
                }
                else
                {
                    PluginLog.Write(
                        $"[LteData][FieldGroupLinks] 活动地组 \"{fieldPrefix}\" 解析失败，fieldIndex=0，field链接将全部跳过"
                    );
                }
            }

            fieldList = fieldList
                .Skip(2)
                .Where(str => !string.IsNullOrEmpty(str)) // 过滤空字符串和null
                .Where(str => str != "0") // 过滤空字符串和null
                .OrderBy(str => str) // 按字母顺序排序
                .ToList();

            if (fieldIndex > 0)
            {
                for (int i = 0; i < fieldList.Count; i++)
                {
                    var fieldValue = fieldList[i];
                    if (double.TryParse(fieldValue, out double fieldValueDouble))
                    {
                        fieldLinks +=
                            "{4,"
                            + (fieldIndex + fieldValueDouble + 1).ToString(
                                CultureInfo.InvariantCulture
                            )
                            + "},";
                    }
                }
            }
        }

        return fieldLinks;
    }
    #endregion

    #region LTE任务数据计算

    public static void FirstCopyTaskValue(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(FirstCopyTaskValue),
            () => CopyTaskValueCore(isFirst: true)
        );

    public static void UpdateCopyTaskValue(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(UpdateCopyTaskValue),
            () => CopyTaskValueCore(isFirst: false)
        );

    private static void CopyTaskValueCore(bool isFirst)
    {
        var sheetName = "LTE【任务】";
        var colIndexArray = PubMetToExcel.ReadExcelDataC(sheetName, 0, 0, 1, 1);
        double activtiyId = colIndexArray?[0, 0] is double d1 ? d1 : 0;

        object[,] copyTaskArray = FilterRepeatValue(
            ActivityDataMinIndex,
            ActivityDataMaxIndex,
            false,
            false
        );

        var taskList = RequireListObject("LTE【任务】", "LTE【任务】");
        if (taskList == null)
            return;

        var baseList = RequireListObject("LTE【基础】", "LTE【基础】");
        if (baseList == null)
            return;
        object[,] baseArray = GetBodyRange(baseList, "LTE【基础】");
        if (baseArray == null)
            return;
        if (baseList.HeaderRowRange?.Value2 is not object[,] baseTitleArray)
        {
            MessageBox.Show("LTE【基础】标题行读取失败");
            return;
        }

        var fieldGroupList = RequireListObject("#道具信息", "道具信息");
        if (fieldGroupList == null)
            return;
        object[,] fieldGroupArray = GetBodyRange(fieldGroupList, "道具信息");
        if (fieldGroupArray == null)
            return;

        var listObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#各类枚举"
        );
        object[,] taskDataTypeArray = GetTableData(listObjectsDic, "#各类枚举", "任务类型");
        if (taskDataTypeArray == null)
            return;
        object[,] dataTypeArray = GetTableData(listObjectsDic, "#各类枚举", "数据类型");
        if (dataTypeArray == null)
            return;

        var findRanklistObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#寻找优先级"
        );
        var findRankDataArray = GetTableData(findRanklistObjectsDic, "#寻找优先级", "寻找方案");
        if (findRankDataArray == null)
            return;

        var copyTaskData = TaskData(
            copyTaskArray,
            taskDataTypeArray,
            baseArray,
            activtiyId,
            fieldGroupArray,
            baseTitleArray,
            findRankDataArray,
            dataTypeArray
        );
        copyTaskArray = copyTaskData.taskArray;
        var errorTypeList = copyTaskData.errorTypeList;
        if (errorTypeList.Count != 0)
        {
            var errorTypeListOnly = new HashSet<string>(errorTypeList);
            MessageBox.Show($"任务数据中存在以下错误类型：{string.Join(",", errorTypeListOnly)}");
        }

        if (isFirst)
        {
            var rowMax = copyTaskArray.GetLength(0);
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                1,
                10000,
                TaskDataStartCol,
                TaskDataEndCol,
                null
            );
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                1,
                rowMax,
                TaskDataStartCol,
                TaskDataEndCol,
                copyTaskArray
            );
            object[,] newTitleArray = PubMetToExcel.ConvertList1ToArrayRow(_taskTitleArray);
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                0,
                0,
                TaskDataStartCol,
                TaskDataEndCol,
                newTitleArray
            );
            taskList.Resize(taskList.Range.Resize[rowMax + 1, taskList.Range.Columns.Count]);
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                1,
                10000,
                TaskDataTagCol,
                TaskDataTagCol,
                null
            );
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
        else
        {
            WriteDymaicData(
                copyTaskArray,
                taskList,
                sheetName,
                TaskDataStartCol,
                TaskDataEndCol,
                TaskDataTagCol
            );
        }
    }

    private static bool CheckLandmarkIdExists(
        Dictionary<string, List<string>> baseDic,
        string id,
        string rank
    )
    {
        if (baseDic.ContainsKey(id))
            return true;
        var sourceId = Convert.ToDouble(id) + Convert.ToDouble(rank) - 1;
        MessageBox.Show($"{sourceId}物品的级别标错了");
        return false;
    }

    //原始数据改造
    private static (object[,] taskArray, List<string> errorTypeList) TaskData(
        object[,] copyTaskArray,
        object[,] taskDataTypeArray,
        object[,] baseArray,
        double activtiyId,
        object[,] fieldGroupArray,
        object[,] baseTitleArray,
        object[,] findRankDataArray,
        object[,] dataTypeArray
    )
    {
        var baseDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(baseArray);
        var taskDataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(taskDataTypeArray);
        var fieldGroupDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(fieldGroupArray);
        var findRankDataDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(findRankDataArray);

        var baseTitleDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(baseTitleArray);

        var dataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(dataTypeArray);

        var titleList = baseTitleDic["数据编号"];

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
            var taskTimeLimit = copyTaskArray[i, 6]?.ToString() ?? String.Empty;

            var taskSubId = copyTaskArray[i, 7]?.ToString() ?? String.Empty;
            var taskSubDes = copyTaskArray[i, 8]?.ToString() ?? String.Empty;
            var taskSubTypeName = copyTaskArray[i, 9]?.ToString() ?? String.Empty;
            var taskSubDialogId = copyTaskArray[i, 10]?.ToString() ?? String.Empty;
            var taskSubTagetName = copyTaskArray[i, 11]?.ToString() ?? String.Empty;
            var taskSubTimeLimit = copyTaskArray[i, 12]?.ToString() ?? String.Empty;

            //改造数据

            var fixMainData = FixTaskData(
                taskTypeName,
                taskDialogId,
                taskTagetName,
                activtiyId,
                taskDataTypeDic,
                baseDic,
                titleList
            );
            if (fixMainData is null)
                continue;
            string taskTypeId = fixMainData[0];
            string taskTagetId = fixMainData[1];
            taskDialogId = fixMainData[2];
            string taskTagetRank = fixMainData[3];

            var fixSubData = FixTaskData(
                taskSubTypeName,
                taskSubDialogId,
                taskSubTagetName,
                activtiyId,
                taskDataTypeDic,
                baseDic,
                titleList
            );
            if (fixSubData is null)
                continue;
            string taskSubTypeId = fixSubData[0];
            string taskSubTagetId = fixSubData[1];
            taskSubDialogId = fixSubData[2];
            string taskSubTagetRank = fixSubData[3];

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
                if (taskTagetId == null)
                {
                    continue;
                }

                // 检查地标类ID经过修正后是否真实存在
                if (!CheckLandmarkIdExists(baseDic, taskTagetId, taskTagetRank))
                    continue;

                var taskTargetMapName = baseDic[taskTagetId][titleList.IndexOf("首次出现")];
                taskColDataList.Add(taskTargetMapName);

                //目标寻找关系
                var findTargetType = baseDic[taskTagetId][titleList.IndexOf("寻找类型")];
                var findTargetDetailType = baseDic[taskTagetId][titleList.IndexOf("寻找细类")];

                var findLinksGroup = FindLinks(
                    findTargetDetailType,
                    findTargetType,
                    taskTagetId,
                    baseDic,
                    out _,
                    fieldGroupDic,
                    titleList,
                    findRankDataDic,
                    dataTypeDic
                );
                var findLinks = findLinksGroup.findLinks;
                var findLinks31 = findLinksGroup.findLinks31;

                taskTargetMapName = MapPrefix(taskTargetMapName);
                var match = _digitsRegex.Match(taskTargetMapName);
                var taskTargetMapId = match.Success ? match.Value : "0";

                if (double.TryParse(taskTargetMapId, out double taskTargetMapIdDouble))
                {
                    taskTargetMapId = (taskTargetMapIdDouble + activtiyId).ToString(
                        CultureInfo.InvariantCulture
                    );
                }

                var taskMapEntrance =
                    taskTargetMapId != "0"
                        ? FindLinkMapEntrance + taskTargetMapId + FindLinkEndNormal
                        : string.Empty;
                findLinks = findLinks + findLinks31 + taskMapEntrance;
                taskColDataList.Add(findLinks);

                // 限时任务数据
                taskColDataList.Add(taskTimeLimit);

                taskColDataList.Add(taskTagetRank);
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
                    taskSubNextId = copyTaskArray[i + 1, 7]?.ToString() ?? string.Empty;
                }
                taskSubColDataList.Add(taskSubNextId);

                // 检查地标类ID经过修正后是否真实存在
                if (!CheckLandmarkIdExists(baseDic, taskSubTagetId, taskSubTagetRank))
                    continue;

                //目标所在地图
                var taskSubTargetMapName = baseDic[taskSubTagetId][titleList.IndexOf("首次出现")];
                taskSubColDataList.Add(taskSubTargetMapName);

                //目标寻找关系
                var findSubTargetType = baseDic[taskSubTagetId][titleList.IndexOf("寻找类型")];
                var findSubTargetDetailType = baseDic[taskSubTagetId][
                    titleList.IndexOf("寻找细类")
                ];

                var findSubLinksGroup = FindLinks(
                    findSubTargetDetailType,
                    findSubTargetType,
                    taskSubTagetId,
                    baseDic,
                    out _,
                    fieldGroupDic,
                    titleList,
                    findRankDataDic,
                    dataTypeDic
                );
                var findSubLinks = findSubLinksGroup.findLinks;
                var findSubLinks31 = findSubLinksGroup.findLinks31;

                taskSubTargetMapName = MapPrefix(taskSubTargetMapName);
                var match = _digitsRegex.Match(taskSubTargetMapName);
                var taskSubTargetMapId = match.Success ? match.Value : "0";

                if (double.TryParse(taskSubTargetMapId, out double taskSubTargetMapIdDouble))
                {
                    taskSubTargetMapId = (taskSubTargetMapIdDouble + activtiyId).ToString(
                        CultureInfo.InvariantCulture
                    );
                }
                findSubLinks =
                    findSubLinks
                    + findSubLinks31
                    + FindLinkMapEntrance
                    + taskSubTargetMapId
                    + FindLinkEndNormal;
                taskSubColDataList.Add(findSubLinks);

                // 限时任务数据
                taskSubColDataList.Add(taskSubTimeLimit);

                taskSubColDataList.Add(taskSubTagetRank);
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

    internal static List<string> FixTaskData(
        string taskTypeName,
        string taskDialogId,
        string taskTagetName,
        double activtiyId,
        Dictionary<string, List<string>> taskDataTypeDic,
        Dictionary<string, List<string>> baseDic,
        List<string> titleList
    )
    {
        var fixData = new List<string>();

        string taskTypeId = string.Empty;
        string taskTagetId = string.Empty;
        string taskTagetRank = string.Empty;

        if (taskTypeName != string.Empty)
        {
            if (!taskDataTypeDic.ContainsKey(taskTypeName))
            {
                MessageBox.Show($"任务类型{taskTypeName}不存在");
                return null;
            }
        }
        if (taskTypeName != string.Empty)
        {
            taskTypeId = taskDataTypeDic[taskTypeName][1] ?? string.Empty;
        }
        if (taskTagetName != string.Empty)
        {
            taskTagetId =
                baseDic
                    .FirstOrDefault(kv =>
                        kv.Value.Count > BaseDicColName && kv.Value[BaseDicColName] == taskTagetName
                    )
                    .Key
                ?? string.Empty;
            if (taskTagetId == string.Empty)
                taskTagetId =
                    baseDic
                        .FirstOrDefault(kv =>
                            kv.Value.Count > BaseDicColOnlyName
                            && kv.Value[BaseDicColOnlyName] == taskTagetName
                        )
                        .Key
                    ?? string.Empty;
            if (taskTagetId == string.Empty)
            {
                MessageBox.Show($"目标名\"{taskTagetName}\"在基础数据中不存在");
                return null;
            }
            var taskTagetType = baseDic.ContainsKey(taskTagetId)
                ? baseDic[taskTagetId][titleList.IndexOf("类型")]
                : string.Empty;

            PluginLog.Write($"[LteData][FixTaskData] 原始目标ID：{taskTagetId}");

            if (
                taskTagetType.Contains("地标-3")
                || taskTagetType == "地标"
                || taskTagetType == "地标-结果"
                || taskTagetType == "地标-结束"
            )
            {
                taskTagetRank = baseDic[taskTagetId][titleList.IndexOf("级别")];
                taskTagetId = (
                    Convert.ToDouble(taskTagetId) - Convert.ToDouble(taskTagetRank) + 1
                ).ToString(CultureInfo.InvariantCulture);

                PluginLog.Write($"[LteData][FixTaskData] 改造目标ID：{taskTagetId}");
            }
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
        fixData.Add(taskTagetId);
        fixData.Add(taskDialogId);

        if (string.IsNullOrEmpty(taskTagetRank))
        {
            taskTagetRank = "1";
        }

        fixData.Add(taskTagetRank);

        return fixData;
    }
    #endregion

    #region LTE地组数据计算
    public static void GroundDataSim(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(GroundDataSim),
            () =>
            {
                var selectedRange = AppServices.App.Selection;
                var targetWorkbookName = "地组工具.xlsx";
                var selectedSheet = AppServices.App.ActiveSheet;
                string targetSheetName = selectedSheet.Name;

                Workbook targetWorkbook = null;

                if (selectedRange == null)
                    throw new InvalidOperationException("没有选中的单元格");

                // 3. 查找已打开的目标工作簿（按名称匹配）
                foreach (Workbook workbook in AppServices.App.Workbooks)
                {
                    if (
                        workbook.Name.Equals(targetWorkbookName, StringComparison.OrdinalIgnoreCase)
                    )
                    {
                        targetWorkbook = workbook;
                        break;
                    }
                }

                if (targetWorkbook == null)
                    throw new FileNotFoundException($"工作簿 '{targetWorkbookName}' 未打开");

                // 4. 获取目标工作表
                var targetSheet = targetWorkbook.Sheets[targetSheetName] as Worksheet;
                if (targetSheet == null)
                    throw new ArgumentException($"目标工作表 '{targetSheetName}' 不存在");

                // 5. 写入值和背景色到目标位置
                var targetRange = targetSheet.Range[selectedRange.Address];

                //// 同步值
                //targetRange.Value = "1";

                //// 同步背景色（使用RGB颜色）
                //targetRange.Interior.Color = 0xFFFF00;

                //// 颜色和数值怎么同步过去？要兼容target已经填的值和颜色

                var targetSheetNameSplit = targetSheetName.Split("_");
                var mapIndex = targetSheetNameSplit.Last();

                // 组成导出的字符串并复制到剪切板
                try
                {
                    // 使用字典存储每个selectedValue对应的所有targetValue
                    var valueMapping = new Dictionary<string, HashSet<string>>();

                    // 遍历所有单元格
                    for (int i = 1; i <= selectedRange.Rows.Count; i++)
                    {
                        for (int j = 1; j <= selectedRange.Columns.Count; j++)
                        {
                            try
                            {
                                // 获取源单元格的值
                                var cell = selectedRange.Cells[i, j];
                                var selectedValue = cell.Value?.ToString()?.Trim() ?? "";

                                // 获取目标区域对应单元格的值
                                var targetCell = targetRange.Cells[i, j];
                                var targetValue = targetCell.Value?.ToString()?.Trim() ?? "";

                                // 添加到字典
                                if (!string.IsNullOrEmpty(selectedValue))
                                {
                                    if (!valueMapping.ContainsKey(selectedValue))
                                    {
                                        valueMapping[selectedValue] = new HashSet<string>();
                                    }
                                    valueMapping[selectedValue].Add(targetValue);
                                }
                            }
                            catch (Exception cellEx)
                            {
                                LogDisplay.RecordLine(
                                    $"处理单元格[{i},{j}]时出错: {cellEx.Message}"
                                );
                                PluginLog.Write($"[LTEData] 单元格异常 [{i},{j}]: {cellEx}");
                            }
                        }
                    }

                    // 构建输出字符串
                    var valueString = new StringBuilder();

                    foreach (var kvp in valueMapping)
                    {
                        string selectedValue = kvp.Key;
                        var targetValues = kvp.Value;

                        if (targetValues.Count == 1)
                        {
                            // 只有一个对应值
                            valueString.AppendLine(
                                $"{selectedValue}\t{mapIndex}\t{targetValues.First()}"
                            );
                        }
                        else
                        {
                            // 多个对应值，用逗号分隔
                            string combinedTargetValues = string.Join("\t", targetValues);
                            valueString.AppendLine(
                                $"{selectedValue}\t{mapIndex}\t{combinedTargetValues}"
                            );
                        }
                    }

                    // 复制到剪切板
                    if (valueString.Length > 0)
                    {
                        ClipboardHelper.SetTextSafe(valueString.ToString());
                    }
                    else
                    {
                        MessageBox.Show(
                            "没有有效数据可复制",
                            "提示",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                    }
                }
                catch (Exception ex)
                {
                    PluginLog.Write($"[LTEData] 复制剪切板失败: {ex}");
                    MessageBox.Show(
                        $"复制到剪切板失败: {ex.Message}",
                        "错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        );

    public static void FirstCopyFieldValue(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(FirstCopyFieldValue),
            () => CopyFieldValueCore(isFirst: true)
        );

    public static void UpdateCopyFieldValue(CommandBarButton ctrl, ref bool cancelDefault) =>
        RunLteCommand(
            ctrl,
            ref cancelDefault,
            nameof(UpdateCopyFieldValue),
            () => CopyFieldValueCore(isFirst: false)
        );

    private static void CopyFieldValueCore(bool isFirst)
    {
        var sheetName = "LTE【地组】";
        var colIndexArray = PubMetToExcel.ReadExcelDataC(sheetName, 0, 0, 1, 1);
        double activtiyId = colIndexArray?[0, 0] is double d2 ? d2 : 0;

        object[,] copyFieldArray = FilterRepeatValue(
            ActivityDataMinIndex,
            ActivityDataMaxIndex,
            false,
            false
        );

        var fieldList = RequireListObject("LTE【地组】", "LTE【地组】");
        if (fieldList == null)
            return;

        var baseList = RequireListObject("LTE【基础】", "LTE【基础】");
        if (baseList == null)
            return;
        object[,] baseArray = GetBodyRange(baseList, "LTE【基础】");
        if (baseArray == null)
            return;
        if (baseList.HeaderRowRange?.Value2 is not object[,] baseTitleArray)
        {
            MessageBox.Show("LTE【基础】标题行读取失败");
            return;
        }

        var fieldGroupList = RequireListObject("#道具信息", "道具信息");
        if (fieldGroupList == null)
            return;
        object[,] fieldGroupArray = GetBodyRange(fieldGroupList, "道具信息");
        if (fieldGroupArray == null)
            return;

        var listObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#各类枚举"
        );
        object[,] fieldDataTypeArray = GetTableData(listObjectsDic, "#各类枚举", "地组类型");
        if (fieldDataTypeArray == null)
            return;
        object[,] dataTypeArray = GetTableData(listObjectsDic, "#各类枚举", "数据类型");
        if (dataTypeArray == null)
            return;

        var findRanklistObjectsDic = PubMetToExcel.GetExcelListObjects(
            WkPath,
            "#【A-LTE】数值大纲.xlsx",
            "#寻找优先级"
        );
        var findRankDataArray = GetTableData(findRanklistObjectsDic, "#寻找优先级", "寻找方案");
        if (findRankDataArray == null)
            return;

        var copyFiledData = FiledData(
            copyFieldArray,
            fieldDataTypeArray,
            baseArray,
            activtiyId,
            fieldGroupArray,
            baseTitleArray,
            findRankDataArray,
            dataTypeArray
        );
        copyFieldArray = copyFiledData.fieldArray;
        var errorTypeList = copyFiledData.errorTypeList;
        if (errorTypeList.Count != 0)
        {
            var errorTypeListOnly = new HashSet<string>(errorTypeList);
            MessageBox.Show($"地组数据中存在以下错误类型：{string.Join(",", errorTypeListOnly)}");
        }

        if (isFirst)
        {
            var rowMax = copyFieldArray.GetLength(0);
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                1,
                10000,
                FieldDataStartCol,
                FieldDataEndCol,
                null
            );
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                1,
                rowMax,
                FieldDataStartCol,
                FieldDataEndCol,
                copyFieldArray
            );
            fieldList.Resize(fieldList.Range.Resize[rowMax + 1, fieldList.Range.Columns.Count]);
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                1,
                10000,
                FieldDataTagCol,
                FieldDataTagCol,
                null
            );
            object[,] writeArray = new object[rowMax, 1];
            for (int i = 0; i < rowMax; i++)
                writeArray[i, 0] = "+";
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                1,
                rowMax,
                FieldDataTagCol,
                FieldDataTagCol,
                writeArray
            );
        }
        else
        {
            WriteDymaicData(
                copyFieldArray,
                fieldList,
                sheetName,
                FieldDataStartCol,
                FieldDataEndCol,
                FieldDataTagCol
            );
        }
    }

    //原始数据改造
    private static (object[,] fieldArray, List<string> errorTypeList) FiledData(
        object[,] copyFieldArray,
        object[,] filedDataTypeArray,
        object[,] baseArray,
        double activtiyId,
        object[,] fieldGroupArray,
        object[,] baseTitleArray,
        object[,] findRankDataArray,
        object[,] dataTypeArray
    )
    {
        var baseDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(baseArray);
        var fieldDataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(filedDataTypeArray);
        var fieldGroupDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(fieldGroupArray);
        var findRankDataDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(findRankDataArray);

        var baseTitleDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(baseTitleArray);

        var dataTypeDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey(dataTypeArray);

        var titleList = baseTitleDic["数据编号"];

        var fieldFieldArrayCount = copyFieldArray.GetLength(0);
        var fieldList = new List<List<string>>();

        var errorTypeList = new List<string>();

        var fieldIdList = new List<string>();
        for (int i = 1; i <= fieldFieldArrayCount; i++)
        {
            var fieldId = copyFieldArray[i, 1]?.ToString() ?? String.Empty;
            if (fieldId != string.Empty)
            {
                fieldIdList.Add(fieldId);
            }
        }

        for (int i = 1; i <= fieldFieldArrayCount; i++)
        {
            var fieldColDataList = new List<string>();

            var fieldId = copyFieldArray[i, 1]?.ToString() ?? String.Empty;

            if (fieldId == string.Empty)
            {
                break;
            }

            var matchFieldIdEnd = fieldId.Substring(fieldId.Length - 1, 1);
            var fieldCount = string.Empty;
            if (matchFieldIdEnd == "1")
            {
                var matchFieldId = fieldId.Substring(0, fieldId.Length - 2);
                var matchCount = fieldIdList.Count(id => id.StartsWith(matchFieldId));

                if (matchCount > 1)
                {
                    if (double.TryParse(fieldId, out double fieldIdDouble))
                    {
                        for (int j = 1; j < matchCount; j++)
                        {
                            fieldCount += (fieldIdDouble + j) + ",";
                        }
                    }
                }
            }
            if (fieldCount != string.Empty)
            {
                fieldCount = fieldCount.Substring(0, fieldCount.Length - 1);
            }

            var fieldDes = copyFieldArray[i, 2]?.ToString() ?? String.Empty;
            var fieldType = copyFieldArray[i, 3]?.ToString() ?? String.Empty;

            if (fieldType != string.Empty)
            {
                if (!fieldDataTypeDic.ContainsKey(fieldType))
                {
                    MessageBox.Show($"地组类型{fieldType}不存在");
                    return (null, null);
                }
            }

            var fieldConditonTarget = copyFieldArray[i, 8]?.ToString() ?? String.Empty;
            var fieldConditonTargetRank = copyFieldArray[i, 5]?.ToString() ?? String.Empty;
            var fieldConditonTargetType = copyFieldArray[i, 9]?.ToString() ?? String.Empty;

            var fieldCost = copyFieldArray[i, 11]?.ToString() ?? String.Empty;

            //改造数据
            string fieldConditon = string.Empty;
            string fieldFindId = string.Empty;

            string findLinks = String.Empty;

            if (fieldConditonTarget != String.Empty)
            {
                var fieldFix = FixFieldData(
                    fieldConditonTarget,
                    fieldConditonTargetRank,
                    fieldConditonTargetType,
                    baseDic
                );
                fieldConditon = fieldFix.fixData;
                fieldFindId = fieldFix.findData;
                var fieldConditonTargetId = fieldFix.fieldConditonTargetId;
                if (
                    !string.IsNullOrEmpty(fieldConditonTargetId)
                    && baseDic.ContainsKey(fieldConditonTargetId)
                )
                {
                    //目标寻找关系
                    var findTargetType = SafeGet(
                        baseDic,
                        fieldConditonTargetId,
                        titleList.IndexOf("寻找类型")
                    );
                    var findTargetDetailType = SafeGet(
                        baseDic,
                        fieldConditonTargetId,
                        titleList.IndexOf("寻找细类")
                    );

                    var findLinksGroup = FindLinks(
                        findTargetDetailType,
                        findTargetType,
                        fieldConditonTargetId,
                        baseDic,
                        out _,
                        fieldGroupDic,
                        titleList,
                        findRankDataDic,
                        dataTypeDic
                    );
                    findLinks = findLinksGroup.findLinks;
                    var findLinks31 = findLinksGroup.findLinks31;

                    var taskTargetMapName = SafeGet(
                        baseDic,
                        fieldConditonTargetId,
                        titleList.IndexOf("首次出现")
                    );
                    taskTargetMapName = MapPrefix(taskTargetMapName);
                    var match = _digitsRegex.Match(taskTargetMapName);
                    var taskTargetMapId = match.Success ? match.Value : "0";

                    if (double.TryParse(taskTargetMapId, out double taskTargetMapIdDouble))
                    {
                        taskTargetMapId = (taskTargetMapIdDouble + activtiyId).ToString(
                            CultureInfo.InvariantCulture
                        );
                    }
                    var fieldMapEntrance =
                        taskTargetMapId != "0"
                            ? FindLinkMapEntrance + taskTargetMapId + FindLinkEndNormal
                            : string.Empty;
                    findLinks = findLinks + findLinks31 + fieldMapEntrance;
                }
            }

            // 消耗寻找
            var fieldCostTarget = copyFieldArray[i, 10]?.ToString() ?? String.Empty;
            string fieldFindId2 =
                baseDic
                    .FirstOrDefault(kv =>
                        kv.Value.Count > BaseDicColName
                        && kv.Value[BaseDicColName] == fieldCostTarget
                    )
                    .Key
                ?? baseDic
                    .FirstOrDefault(kv =>
                        kv.Value.Count > BaseDicColOnlyName
                        && kv.Value[BaseDicColOnlyName] == fieldCostTarget
                    )
                    .Key;

            string findLinks2 = string.Empty;

            if (
                fieldFindId2 != null
                && fieldFindId2 != fieldFindId
                && fieldCostTarget != string.Empty
                && baseDic.ContainsKey(fieldFindId2)
            )
            {
                var findTarget2Type = SafeGet(baseDic, fieldFindId2, titleList.IndexOf("寻找类型"));
                var findTarget2DetailType = SafeGet(
                    baseDic,
                    fieldFindId2,
                    titleList.IndexOf("寻找细类")
                );

                var findLinks2Group = FindLinks(
                    findTarget2DetailType,
                    findTarget2Type,
                    fieldFindId2,
                    baseDic,
                    out _,
                    fieldGroupDic,
                    titleList,
                    findRankDataDic,
                    dataTypeDic
                );
                findLinks2 = findLinks2Group.findLinks;
                var findLinks231 = findLinks2Group.findLinks31;

                var taskTarget2MapName = baseDic[fieldFindId2][titleList.IndexOf("首次出现")];
                taskTarget2MapName = MapPrefix(taskTarget2MapName);
                var match2 = _digitsRegex.Match(taskTarget2MapName);
                var taskTarget2MapId = match2.Success ? match2.Value : "0";

                if (double.TryParse(taskTarget2MapId, out double taskTarget2MapIdDouble))
                {
                    taskTarget2MapId = (taskTarget2MapIdDouble + activtiyId).ToString(
                        CultureInfo.InvariantCulture
                    );
                }
                findLinks2 =
                    findLinks2
                    + findLinks231
                    + FindLinkMapEntrance
                    + taskTarget2MapId
                    + FindLinkEndNormal;
            }
            else
            {
                fieldFindId2 = string.Empty;
            }

            if (fieldId != string.Empty)
            {
                fieldColDataList.Add(fieldId);
                fieldColDataList.Add(fieldType);
                fieldColDataList.Add(fieldDes);
                fieldColDataList.Add(fieldConditon);
                fieldColDataList.Add(fieldCost);
                fieldColDataList.Add(fieldFindId);
                fieldColDataList.Add(findLinks);
                fieldColDataList.Add(fieldFindId2);
                fieldColDataList.Add(findLinks2);
                fieldColDataList.Add(fieldCount);
            }

            if (fieldColDataList.Count != 0)
            {
                fieldList.Add(fieldColDataList);
            }
        }
        var fieldArray = PubMetToExcel.ConvertListToArray(fieldList);
        return (fieldArray, errorTypeList);
    }

    internal static (string fixData, string findData, string fieldConditonTargetId) FixFieldData(
        string fieldConditonTarget,
        string fieldConditonTargetRank,
        string fieldConditonTargetType,
        Dictionary<string, List<string>> baseDic
    )
    {
        var fixData = string.Empty;
        var findData = string.Empty;
        var fixIcon = string.Empty;

        var fieldConditonTargetId =
            baseDic
                .FirstOrDefault(kv =>
                    kv.Value.Count > BaseDicColName
                    && kv.Value[BaseDicColName] == fieldConditonTarget
                )
                .Key
            ?? baseDic
                .FirstOrDefault(kv =>
                    kv.Value.Count > BaseDicColOnlyName
                    && kv.Value[BaseDicColOnlyName] == fieldConditonTarget
                )
                .Key;

        if (fieldConditonTargetId == null)
        {
            MessageBox.Show($"{fieldConditonTarget}:找不到ID，检查【基础】表");
            return (string.Empty, string.Empty, string.Empty);
        }

        var fieldConditonTargetName = baseDic[fieldConditonTargetId][BaseDicColPackage];
        var fieldConditonTargetLast = baseDic
            .LastOrDefault(kv =>
                kv.Value.Count > BaseDicColPackage
                && kv.Value[BaseDicColPackage] == fieldConditonTargetName
            )
            .Key;

        var fieldConditonTargetPic = fieldConditonTargetId.Substring(0, 8) + "00";

        if (fieldConditonTargetRank != string.Empty || fieldConditonTarget != String.Empty)
        {
            if (double.TryParse(fieldConditonTargetRank, out double fieldConditonTargetRankDouble))
            {
                if (double.TryParse(fieldConditonTargetId, out double fieldConditonTargetIdDouble))
                {
                    fixIcon = Convert.ToString(
                        fieldConditonTargetIdDouble + fieldConditonTargetRankDouble,
                        CultureInfo.InvariantCulture
                    );
                }
            }

            if (fieldConditonTargetType.Contains("地标"))
            {
                fixData = $"[[14,{fieldConditonTargetId},{fieldConditonTargetRank},{fixIcon}]]";
                findData = fieldConditonTargetId;
            }
            else if (fieldConditonTargetType.Contains("链"))
            {
                fixData = $"[[7,{fieldConditonTargetId},{fieldConditonTargetPic}]]";
            }
            else if (fieldConditonTargetType.StartsWith("修"))
            {
                fixData = $"[[8,{fieldConditonTargetId},{fieldConditonTargetLast}]]";
            }
            else if (fieldConditonTargetType.StartsWith("兑-篝火"))
            {
                fixData = $"[[12,{fieldConditonTargetId},-1]]";
            }
            else
            {
                fixData = $"[[7,{fieldConditonTargetId}]]";
                findData = fieldConditonTargetId;
            }
        }
        return (fixData, findData, fieldConditonTargetId);
    }

    #endregion
}
