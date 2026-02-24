using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using MiniExcelLibs;
using NLua;
using NumDesTools.Config;
using NumDesTools.UI;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Clipboard = System.Windows.Forms.Clipboard;
using Match = System.Text.RegularExpressions.Match;
using MessageBox = System.Windows.MessageBox;
using Process = System.Diagnostics.Process;

// ReSharper disable All

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类调用的具体业务逻辑
/// </summary>
public static class PubMetToExcelFunc
{
    private static readonly dynamic Wk = NumDesAddIn.App.ActiveWorkbook;

    private static readonly string WkPath = Wk.Path;

    public static void ExcelDataSearchAndMerge(string searchValue)
    {
        string[] ignoreFileNames = ["#", "副本"];
        var rootPath = Path.GetDirectoryName(Path.GetDirectoryName(WkPath));
        var fileList = new List<string>()
        {
            rootPath + @"\Excels\Tables\",
            rootPath + @"\Excels\Localizations\",
            rootPath + @"\Excels\UIs\"
        };
        var files = PubMetToExcel.PathExcelFileCollect(fileList, "*.xlsx", ignoreFileNames);
        var findValueList = new List<(string, string, int, int, string, string)>();
        Parallel.ForEach(
            files,
            file =>
            {
                var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(file);
                var findValue = PubMetToExcel.FindDataInDataTable(file, dataTable, searchValue);
                if (findValue.Count > 0)
                    findValueList = findValueList.Concat(findValue).ToList();
            }
        );
        dynamic tempWorkbook;
        try
        {
            tempWorkbook = NumDesAddIn.App.Workbooks.Open(
                rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx"
            );
        }
        catch
        {
            tempWorkbook = NumDesAddIn.App.Workbooks.Add();
            tempWorkbook.SaveAs(rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
        }

        var tempSheet = tempWorkbook.Sheets["Sheet1"];
        var tempDataArray = new string[findValueList.Count, 5];
        for (var i = 0; i < findValueList.Count; i++)
        {
            tempDataArray[i, 0] = findValueList[i].Item1;
            tempDataArray[i, 1] = findValueList[i].Item2;
            tempDataArray[i, 2] =
                PubMetToExcel.ConvertToExcelColumn(findValueList[i].Item4) + findValueList[i].Item3;
            tempDataArray[i, 3] = findValueList[i].Item5;
            tempDataArray[i, 4] = findValueList[i].Item6;
        }

        var tempDataRange = tempSheet.Range[
            tempSheet.Cells[2, 2],
            tempSheet.Cells[2 + tempDataArray.GetLength(0) - 1, 2 + tempDataArray.GetLength(1) - 1]
        ];
        tempDataRange.Value = tempDataArray;
        tempWorkbook.Save();
    }

    public static void RightOpenExcelByActiveCell(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件
        var sheet = NumDesAddIn.App.ActiveSheet;
        var selectCell = NumDesAddIn.App.ActiveCell;
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var workBookName = workBook.Name;
        var workbookPath = workBook.Path;
        workbookPath = Path.GetDirectoryName(workbookPath);

        var selectCellValue = "";
        if (selectCell.Value != null)
            selectCellValue = selectCell.Value.ToString();
        var isMatch = selectCellValue.Contains(".xls");
        if (isMatch)
        {
            string sheetName;
            var cellAddress = "A1";
            if (workBookName.Contains("#合并表格数据缓存"))
            {
                var selectRow = selectCell.Row;
                var selectCol = selectCell.Column;
                sheetName = sheet.Cells[selectRow, selectCol + 1].Value;
                cellAddress = sheet.Cells[selectRow, selectCol + 2].Value;
            }
            else if (selectCellValue.Contains("#") && !selectCellValue.Contains("##"))
            {
                var excelSplit = selectCellValue.Split("#");
                selectCellValue = workbookPath + @"\Tables\" + excelSplit[0];
                sheetName = excelSplit[1];
            }
            else if (selectCellValue.Contains("##"))
            {
                var excelSplit = selectCellValue.Split("##");
                var sharpCount = excelSplit.Length;
                if (selectCellValue.Contains("克朗代克"))
                {
                    selectCellValue =
                        workbookPath + @"\Tables\" + excelSplit[0] + @"\" + excelSplit[1];
                    sheetName = sharpCount == 3 ? excelSplit[2] : "Sheet1";
                }
                else
                {
                    selectCellValue = workbookPath + @"\Tables\" + excelSplit[0];
                    sheetName = excelSplit[1];
                }
            }
            else
            {
                switch (selectCellValue)
                {
                    case "Localizations.xlsx":
                        selectCellValue = workbookPath + @"\Localizations\Localizations.xlsx";
                        break;
                    case "UIConfigs.xlsx":
                        selectCellValue = workbookPath + @"\UIs\UIConfigs.xlsx";
                        break;
                    case "UIItemConfigs.xlsx":
                        selectCellValue = workbookPath + @"\UIs\UIItemConfigs.xlsx";
                        break;
                    default:
                        selectCellValue = workbookPath + @"\Tables\" + selectCellValue;
                        break;
                }

                sheetName = "Sheet1";
            }

            PubMetToExcel.OpenExcelAndSelectCell(selectCellValue, sheetName, cellAddress);
        }
    }

    public static void RightOpenLinkExcelByActiveCell(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件

        var sheet = NumDesAddIn.App.ActiveSheet;
        var selectCell = NumDesAddIn.App.ActiveCell;
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var workBookName = workBook.Name;
        var workbookPath = workBook.Path;
        var sheetName = sheet.Name;

        //根目录下的子文件路径处理
        if (
            workbookPath!.Contains("克朗代克")
            || workbookPath!.Contains("二合")
            || workbookPath!.Contains("工会")
        )
        {
            workBookName = "克朗代克##" + workBookName;
            workbookPath = Directory.GetParent(Path.GetDirectoryName(workbookPath)).FullName;
        }
        else
        {
            workbookPath = Path.GetDirectoryName(workbookPath);
        }

        var selectCellCol = selectCell.Column;
        var keyCell = sheet.Cells[2, selectCellCol];
        var excelPath = workbookPath + @"\Tables";
        var excelName = "#表格关联.xlsx##主副表关联";
        var excelObj = new ExcelDataByEpplus();
        excelObj.GetExcelObj(excelPath, excelName);
        if (excelObj.ErrorList.Count > 0)
            return;
        var sheetTarget = excelObj.Sheet;
        var data = excelObj.ReadToDic(sheetTarget, 6, 5, [7, 9], 2);

        string keyName;
        //过滤单Sheet工作簿
        if (sheetName.Contains("Sheet") || sheetName.Contains("map_proto"))
            keyName = workBookName;
        else
            keyName = workBookName + "##" + sheetName;

        if (data.TryGetValue(keyName, out var valueList))
        {
            //查找所有满足条件的值，然后按顺序遍历文件，找到第一个存在查找ID的表
            var result = valueList
                .Cast<List<string>>()
                .Where(list => list[0] == keyCell.Value.ToString())
                .ToList();
            if (result.Count != 0)
            {
                OpenTargetExcel(
                    result,
                    selectCell,
                    workbookPath,
                    excelPath,
                    excelName,
                    workBookName
                );
            }
            else
            {
                var tips = MessageBox.Show(
                    "字段未关联或没有表格索引,是否打开字段表格编辑？是：打开表编辑；否：模糊查找相似度最高的关联表，可能找不到",
                    "确认",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question
                );
                if (tips == MessageBoxResult.Yes)
                {
                    PubMetToExcel.OpenExcelAndSelectCell(excelPath + @"\#表格关联.xlsx", "主副表关联", "A1");
                }
                //尝试选择最相似的字段关联表
                else
                {
                    var blurData = data.Values.SelectMany(list => list).ToList();
                    // 查找最相似的键并返回对应的值
                    string blurResultText = FindClosestMatch(blurData, keyCell.Value.ToString(), 2);
                    List<List<string>> blurResult =
                    [
                        ["", blurResultText]
                    ];
                    if (blurResultText == null)
                        return;
                    OpenTargetExcel(
                        blurResult,
                        selectCell,
                        workbookPath,
                        excelPath,
                        excelName,
                        workBookName
                    );
                }
            }
        }
        else
        {
            var tips = MessageBox.Show(
                "字段未关联或没有表格索引,是否打开字段表格编辑？是：打开表编辑；否：模糊查找相似度最高的关联表，可能找不到",
                "确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question
            );
            if (tips == MessageBoxResult.Yes)
            {
                PubMetToExcel.OpenExcelAndSelectCell(excelPath + @"\#表格关联.xlsx", "主副表关联", "A1");
            }
            //尝试选择最相似的字段关联表
            else
            {
                var blurData = data.Values.SelectMany(list => list).ToList();
                // 查找最相似的键并返回对应的值
                string blurResultText = FindClosestMatch(blurData, keyCell.Value.ToString(), 2);
                List<List<string>> blurResult =
                [
                    ["", blurResultText]
                ];
                if (blurResultText == null)
                    return;
                OpenTargetExcel(
                    blurResult,
                    selectCell,
                    workbookPath,
                    excelPath,
                    excelName,
                    workBookName
                );
            }
        }
    }

    private static void OpenTargetExcel(
        List<List<string>> indexCellValueList,
        Range selectCell,
        string workbookPath,
        string excelPath,
        string excelName,
        string workBookName
    )
    {
        foreach (var wkNameList in indexCellValueList)
        {
            var indexCellValue = wkNameList[1];
            //活动主表ActivityID特殊处理
            if (indexCellValue == "活动编号")
            {
                var excelObj = new ExcelDataByEpplus();
                excelObj.GetExcelObj(excelPath, "#表格关联.xlsx##活动类型枚举");
                if (excelObj.ErrorList.Count > 0)
                    return;
                var sheetTarget = excelObj.Sheet;
                var data = excelObj.ReadToDic(sheetTarget, 6, 7, [7, 8], 2);
                var selectCellCol = selectCell.Column;
                var selectCellRow = selectCell.Row;
                var sheet = NumDesAddIn.App.ActiveSheet;
                var typeCell = sheet.Cells[selectCellRow, selectCellCol - 1];
                string typeValue = typeCell.Value.ToString();
                if (data.TryGetValue(typeValue, out var valueList))
                {
                    var result = valueList
                        .Cast<List<string>>()
                        .FirstOrDefault(list => list[0] == typeValue);

                    if (result != null)
                        indexCellValue = result[1];
                }
            }

            var isMatch = indexCellValue.Contains(".xls");
            var wkName = indexCellValue;
            if (isMatch)
            {
                string openSheetName;
                var selectCellValue = selectCell.Value.ToString();
                if (indexCellValue.Contains("##"))
                {
                    var excelSplit = indexCellValue.Split("##");
                    var sharpCount = excelSplit.Length;
                    if (indexCellValue.Contains("克朗代克"))
                    {
                        indexCellValue =
                            workbookPath + @"\Tables\" + excelSplit[0] + @"\" + excelSplit[1];
                        openSheetName = sharpCount == 3 ? excelSplit[2] : "Sheet1";
                    }
                    else
                    {
                        indexCellValue = workbookPath + @"\Tables\" + excelSplit[0];
                        openSheetName = excelSplit[1];
                    }
                }
                else
                {
                    switch (indexCellValue)
                    {
                        case "Localizations.xlsx":
                            indexCellValue = workbookPath + @"\Localizations\Localizations.xlsx";
                            break;
                        case "UIConfigs.xlsx":
                            indexCellValue = workbookPath + @"\UIs\UIConfigs.xlsx";
                            break;
                        case "UIItemConfigs.xlsx":
                            indexCellValue = workbookPath + @"\UIs\UIItemConfigs.xlsx";
                            break;
                        default:
                            indexCellValue = workbookPath + @"\Tables\" + indexCellValue;
                            break;
                    }

                    openSheetName = "Sheet1";
                }

                var excelLinkObjOpen = new ExcelDataByEpplus();
                excelLinkObjOpen.GetExcelObj(excelPath, excelName);
                if (excelLinkObjOpen.ErrorList.Count > 0)
                    return;
                var sheetLinkOpen = excelLinkObjOpen.Sheet;
                var valueLinkIndex = excelLinkObjOpen.FindFromRow(sheetLinkOpen, 2, workBookName);
                var cellLinkAddress = "A1";
                if (valueLinkIndex != -1)
                    cellLinkAddress = "A" + valueLinkIndex;
                if (!File.Exists(indexCellValue))
                {
                    var tips = MessageBox.Show(
                        "文件[" + indexCellValue + "]不存在，是否打开字段表格编辑？",
                        "确认",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question
                    );
                    if (tips == MessageBoxResult.Yes) { }

                    PubMetToExcel.OpenExcelAndSelectCell(
                        excelPath + @"\#表格关联.xlsx",
                        "主副表关联",
                        cellLinkAddress
                    );
                    return;
                }

                var pattern = @"\d+";
                if (
                    indexCellValue.Contains("Localizations.xlsx")
                    || indexCellValue.Contains("UIConfigs.xlsx")
                    || indexCellValue.Contains("UIItemConfigs.xlsx")
                )
                {
                    pattern = @".*";
                }

                MatchCollection matches = Regex.Matches(selectCellValue, pattern);
                var cellAddress = "A1";
                var excelObjOpen = new ExcelDataByEpplus();

                string excelNameOpen;
                if (wkName.Contains("##"))
                {
                    var isKol = wkName.Substring(wkName.Length - 4, 4);
                    if (isKol == "xlsx")
                    {
                        var excelSplit = wkName.Split("##");
                        wkName = excelSplit[1];
                        excelNameOpen = wkName + "##Sheet1";
                    }
                    else
                    {
                        excelNameOpen = wkName;
                    }
                }
                else
                {
                    excelNameOpen = wkName + "##Sheet1";
                }

                if (indexCellValue.Contains("##"))
                    excelNameOpen = wkName;

                excelObjOpen.GetExcelObj(Path.GetDirectoryName(indexCellValue), excelNameOpen);

                if (excelObjOpen.ErrorList.Count > 0)
                    return;
                var sheetTargetOpen = excelObjOpen.Sheet;
                foreach (var item in matches)
                {
                    var valueIndex = excelObjOpen.FindFromRow(sheetTargetOpen, 2, item.ToString());
                    if (valueIndex != -1)
                    {
                        cellAddress = "A" + valueIndex;
                        break;
                    }
                }

                if (cellAddress == "A1")
                {
                    var tips = MessageBox.Show(
                        "文件[" + indexCellValue + "]不存在查找字段，是否继续",
                        "确认",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question
                    );
                    if (tips == MessageBoxResult.Yes)
                    {
                        continue;
                    }

                    PubMetToExcel.OpenExcelAndSelectCell(
                        indexCellValue,
                        openSheetName,
                        cellAddress
                    );
                    continue;
                }

                PubMetToExcel.OpenExcelAndSelectCell(indexCellValue, openSheetName, cellAddress);
            }
        }
    }

    private static string FindClosestMatch(List<object> listOfObjects, string input, int threshold)
    {
        List<List<string>> listOfLists = new List<List<string>>();

        // 将 List<object> 转换为 List<List<string>>
        foreach (var obj in listOfObjects)
        {
            if (obj is List<string> list && list.Count == 2)
            {
                listOfLists.Add(list);
            }
            else
            {
                throw new InvalidCastException(
                    "The object is not of type List<string> with exactly two elements."
                );
            }
        }

        // 查找最相似的键
        int closestDistance = int.MaxValue;
        string closestValue = null;

        foreach (var list in listOfLists)
        {
            string key = list[0];
            int distance = LevenshteinDistance(input, key);
            if (distance < closestDistance && distance <= threshold)
            {
                closestDistance = distance;
                closestValue = list[1];
            }
        }

        // 返回最相似键对应的值
        return closestValue;
    }

    // 计算两个字符串之间的 Levenshtein 距离
    static int LevenshteinDistance(string s, string t)
    {
        int n = s.Length;
        int m = t.Length;
        int[,] d = new int[n + 1, m + 1];

        if (n == 0)
            return m;
        if (m == 0)
            return n;

        for (int i = 0; i <= n; d[i, 0] = i++) { }

        for (int j = 0; j <= m; d[0, j] = j++) { }

        for (int i = 1; i <= n; i++)
        {
            for (int j = 1; j <= m; j++)
            {
                int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                d[i, j] = Math.Min(
                    Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                    d[i - 1, j - 1] + cost
                );
            }
        }

        return d[n, m];
    }

    public static void OpenBaseLanExcel(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件

        var selectCell = NumDesAddIn.App.ActiveCell;
        var basePath = NumDesAddIn.App.ActiveWorkbook.Path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(basePath));
        newPath = newPath + @"\Excels\Localizations\Localizations.xlsx";
        var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(newPath);
        var findValue = PubMetToExcel.FindDataInDataTable(
            newPath,
            dataTable,
            selectCell.Value.ToString()
        );
        var cellAddress =
            PubMetToExcel.ConvertToExcelColumn(findValue[0].Item4) + findValue[0].Item3;
        PubMetToExcel.OpenExcelAndSelectCell(newPath, "Sheet1", cellAddress);
    }

    public static void OpenMergeLanExcel(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件

        var selectCell = NumDesAddIn.App.ActiveCell;
        var basePath = NumDesAddIn.App.ActiveWorkbook.Path;
        var mergePath = "";
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        if (mergePathList.Count <= 1)
            Process.Start(filePath);
        if (
            mergePathList[0] == ""
            || mergePathList[1] == ""
            || mergePathList[1] == mergePathList[0]
        )
            Process.Start(filePath);
        else
            mergePath = basePath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(mergePath));
        newPath = newPath + @"\Localizations\Localizations.xlsx";
        var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(newPath);
        var findValue = PubMetToExcel.FindDataInDataTable(
            newPath,
            dataTable,
            selectCell.Value.ToString()
        );
        string cellAddress;
        if (findValue.Count == 0)
            cellAddress = "A1";
        else
            cellAddress =
                PubMetToExcel.ConvertToExcelColumn(findValue[0].Item4) + findValue[0].Item3;
        PubMetToExcel.OpenExcelAndSelectCell(newPath, "Sheet1", cellAddress);
    }

    //大富翁种
    public static void AliceBigRicherDfs2(string sheetName)
    {
        var baseName = "大富翁种";
        if (!sheetName.Contains(baseName))
            MessageBox.Show("当前表格不是【大富翁种**】,无法使用大富翁功能");
        var targetRank = PubMetToExcel.ReadExcelDataC(sheetName, 21, 44, 2, 2);
        var targetKey = PubMetToExcel.ReadExcelDataC(sheetName, 13, 15, 2, 3);
        var targetKeySoft = PubMetToExcel.ReadExcelDataC(sheetName, 2, 11, 2, 3);
        var maxRollCell = PubMetToExcel.ReadExcelDataC(sheetName, 18, 18, 2, 2);
        var maxGridLoopCell = PubMetToExcel.ReadExcelDataC(sheetName, 17, 17, 2, 2);
        var maxRankCell = PubMetToExcel.ReadExcelDataC(sheetName, 18, 18, 5, 5);
#pragma warning disable CA1305
        var maxRoll = Convert.ToInt32(maxRollCell[0, 0]);
#pragma warning restore CA1305
#pragma warning disable CA1305
        var maxGridLoop = Convert.ToInt32(maxGridLoopCell[0, 0]);
#pragma warning restore CA1305
#pragma warning disable CA1305
        var maxRankValue = Convert.ToInt32(maxRankCell[0, 0]);
#pragma warning restore CA1305
        var permutations = PubMetToExcel.UniqueRandomMethod(maxRoll, maxRoll * 100000, 6);
        var targetProcess = new Dictionary<int, List<int>>();
        var bpProcess = new Dictionary<int, List<int>>();
        var targetGift = new Dictionary<int, List<int>>();
        var modCountDiv = maxRoll;

        for (var i = 0; i < permutations.Count; i++)
        {
            var targetProcessTemp = new List<int>();
            var bpProcessTemp = new List<int>();
            var targetGiftTemp = new List<int>();
            for (var j = 0; j < maxGridLoop * 24; j++)
            {
                var modCount = (j + 1) % modCountDiv;
                if (modCount == 0)
                    modCount = modCountDiv;
                if (j == 0)
                {
                    targetProcessTemp.Add(permutations[i][0]);
                    bpProcessTemp.Add(permutations[i][0]);
                    var tempValue = targetRank[0, 0];
                    if (tempValue is ExcelEmpty)
                        tempValue = null;
#pragma warning disable CA1305
                    targetGiftTemp.Add(Convert.ToInt32(tempValue));
#pragma warning restore CA1305
                }
                else
                {
                    var targetTemp = targetProcessTemp[j - 1] + permutations[i][modCount - 1];
                    var processTemp = bpProcessTemp[j - 1] + permutations[i][modCount - 1];
                    bpProcessTemp.Add(processTemp);
                    targetTemp %= 24;
                    if (targetTemp == 0)
                        targetTemp = 24;
                    targetProcessTemp.Add(targetTemp);
                    var tempValue = targetRank[targetTemp - 1, 0];
                    if (tempValue is ExcelEmpty)
                        tempValue = null;
#pragma warning disable CA1305
                    var targetTemp2 = targetGiftTemp[j - 1] + Convert.ToInt32(tempValue);
#pragma warning restore CA1305
                    targetGiftTemp.Add(targetTemp2);
                }
            }

            targetProcess[i] = targetProcessTemp;
            bpProcess[i] = bpProcessTemp;
            targetGift[i] = targetGiftTemp;
        }

        var filteredData = targetProcess;
        for (var i = 0; i < targetKey.GetLength(0); i++)
        {
            var rollTimes = targetKey[i, 0];
            var rollGrid = targetKey[i, 1];
            if (!(rollTimes is ExcelEmpty))
            {
#pragma warning disable CA1305
                var colIndex = Convert.ToInt32(rollTimes) - 1;
#pragma warning restore CA1305
#pragma warning disable CA1305
                var colValue = Convert.ToInt32(rollGrid);
#pragma warning restore CA1305
                filteredData = filteredData
                    .Where(entry => entry.Value[colIndex] == colValue)
                    .ToDictionary(entry => entry.Key, entry => entry.Value);
            }
        }

        for (var i = 0; i < targetKeySoft.GetLength(0); i++)
        {
            var softTimes = targetKeySoft[i, 1];
            var softGrid = targetKeySoft[i, 0];
            if (!(softGrid is ExcelEmpty))
#pragma warning disable CA1305
#pragma warning disable CA1305
                filteredData = filteredData
                    .Where(pair =>
                        pair.Value.Take(maxRoll).Count(item => item == Convert.ToInt32(softGrid))
                        == Convert.ToInt32(softTimes)
                    )
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
#pragma warning restore CA1305
#pragma warning restore CA1305
        }

        var filteredDataGift = new Dictionary<int, int>();
        var filteredDataMethod = new List<List<object>>();
        var filteredDataBpProcess = new List<List<object>>();
        foreach (var key in filteredData.Keys)
            filteredDataGift[key] = targetGift[key][maxRoll];
        var modeValue = maxRankValue;
        var filteredDataGiftMode = filteredDataGift.Where(pair => pair.Value == modeValue).ToList();
        var filteredDataGiftList = new List<List<object>>();
        foreach (var kvp in filteredDataGiftMode)
        {
            var key = kvp.Key;
            var value = kvp.Value;
            filteredDataGiftList.Add([value]);
            filteredDataBpProcess.Add([bpProcess[key][maxRoll]]);
            var methodStr = "";
            foreach (var method in permutations[key])
                methodStr += method + ",";
            methodStr = methodStr.Substring(0, methodStr.Length - 1);
            filteredDataMethod.Add([methodStr]);
        }

        var emptyData = new object[65535 - 17 + 1, 6 - 6 + 1];
        PubMetToExcel.WriteExcelDataC(sheetName, 21, 65534, 4, 4, emptyData);
        PubMetToExcel.WriteExcelDataC(sheetName, 21, 65534, 5, 5, emptyData);
        PubMetToExcel.WriteExcelDataC(sheetName, 21, 65534, 6, 6, emptyData);
        if (filteredDataBpProcess.Count == 0)
        {
            var error = new object[1, 1];
            error[0, 0] = "#Error#";
            PubMetToExcel.WriteExcelDataC(sheetName, 21, 21, 4, 4, error);
        }
        else
        {
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                21,
                21 + filteredDataBpProcess.Count - 1,
                6,
                6,
                PubMetToExcel.ConvertListToArray(filteredDataBpProcess)
            );
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                21,
                21 + filteredDataGiftList.Count - 1,
                5,
                5,
                PubMetToExcel.ConvertListToArray(filteredDataGiftList)
            );
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                21,
                21 + filteredDataMethod.Count - 1,
                4,
                4,
                PubMetToExcel.ConvertListToArray(filteredDataMethod)
            );
        }
    }

    //魔瓶验算
    public static void MagicBottleCostSimulate(string sheetName)
    {
        var baseName = "移动魔瓶";
        if (!sheetName.Contains(baseName))
            MessageBox.Show("当前表格不是【移动魔瓶**】,无法使用魔瓶验算");
        var eleCount = PubMetToExcel.ReadExcelDataC(sheetName, 2, 8, 21, 21);
        var simulateCount = PubMetToExcel.ReadExcelDataC(sheetName, 0, 0, 21, 21);
#pragma warning disable CA1305
        var simulateCountMax = Convert.ToInt32(simulateCount[0, 0]);
#pragma warning restore CA1305
        var eleCountMax = eleCount.GetLength(0);
        var filterEleCountMax = new List<int>();
        for (var r = 0; r < eleCountMax; r++)
            filterEleCountMax.Add(0);
        for (var s = 0; s < simulateCountMax; s++)
        for (var r = 0; r < eleCountMax; r++)
        {
            var eleGuessListGroup = new List<List<int>>();
#pragma warning disable CA1305
            var eleNum = Convert.ToInt32(eleCount[r, 0]);
#pragma warning restore CA1305
            var eleList = new List<int>();
            var eleGuessList = new List<int>();
            for (var e = 1; e <= eleNum; e++)
            {
                eleList.Add(e);
                eleGuessList.Add(e);
            }

            var seedTarget = new Random();
            eleList = eleList.OrderBy(_ => seedTarget.Next()).ToList();
            var seedGuess = new Random();
            eleGuessList = eleGuessList.OrderBy(_ => seedGuess.Next()).ToList();
            do
            {
                for (var eleCurrent = eleList.Count - 1; eleCurrent >= 0; eleCurrent--)
                {
                    var ele = eleList[eleCurrent];
                    var eleGuess = eleGuessList[eleCurrent];
                    if (eleGuess == ele)
                    {
                        eleList.RemoveAt(eleCurrent);
                        eleGuessList.RemoveAt(eleCurrent);
                    }

                    filterEleCountMax[r]++;
                }

                eleGuessListGroup.Add(eleGuessList);
                if (eleList.Count > 1)
                {
                    List<int> eleTempList;
                    var seedTemp = new Random();
                    do
                    {
                        eleTempList = eleGuessList.OrderBy(_ => seedTemp.Next()).ToList();
                    } while (eleGuessListGroup.Any(list => list.SequenceEqual(eleTempList)));

                    eleGuessList = eleTempList;
                }
            } while (eleList.Count != 0);
        }

        // ReSharper disable PossibleLossOfFraction
        var filterEleCountMaxObj = filterEleCountMax
            .Select(item => (double)(item / simulateCountMax))
            // ReSharper restore PossibleLossOfFraction
            .Select(simulateValue => new List<object> { simulateValue })
            .ToList();
        var emptyData = new object[7, 7];
        PubMetToExcel.WriteExcelDataC(sheetName, 2, 8, 22, 22, emptyData);
        var emptyData2 = new object[1, 1];
        PubMetToExcel.WriteExcelDataC(sheetName, 0, 0, 23, 23, emptyData2);
        if (filterEleCountMax.Count == 0)
        {
            var error = new object[1, 1];
            error[0, 0] = "#Error#";
            PubMetToExcel.WriteExcelDataC(sheetName, 0, 0, 23, 23, error);
        }
        else
        {
            PubMetToExcel.WriteExcelDataC(
                sheetName,
                2,
                2 + filterEleCountMax.Count - 1,
                22,
                22,
                PubMetToExcel.ConvertListToArray(filterEleCountMaxObj)
            );
        }
    }

    // 抽卡模拟
    public static void PhotoCardRatio(string sheetName)
    {
        var baseName = "相册万能卡";
        if (!sheetName.Contains(baseName))
            MessageBox.Show("当前表格不是【相册万能卡**】,无法使用相册万能卡功能");
        var sourceRatioRange = PubMetToExcel.ReadExcelDataC(sheetName, 1, 100, 1, 2);
        var loopNumRange = PubMetToExcel.ReadExcelDataC(sheetName, 1, 1, 3, 3);

        var rowCount = sourceRatioRange.GetLength(0);

        var sourceRatioDic = new Dictionary<int, double>();
        for (int i = 0; i < rowCount; i++)
        {
            var sourceRatioRow = sourceRatioRange[i, 0]?.ToString();
            if (sourceRatioRow == "ExcelDna.Integration.ExcelEmpty")
            {
                break;
            }
            sourceRatioDic[Convert.ToInt32(sourceRatioRow)] = Convert.ToDouble(
                sourceRatioRange[i, 1]
            );
        }

        //初始化字典
        var simRatioDic = new Dictionary<int, int>();
        int maxKey = sourceRatioDic.Keys.Max();
        for (int i = 1; i <= maxKey; i++)
        {
            simRatioDic[i] = 0;
        }
        //模拟次数
        int loopNum = Convert.ToInt32(loopNumRange[0, 0]);
        for (int i = 0; i < loopNum; i++)
        {
            var getNum = 1;
            bool getIt = false;
            var random = new Random();
            while (!getIt)
            {
                double randSeed = random.NextDouble();
                var getNumValue = GetNearKeyValue(sourceRatioDic, getNum);
                if (randSeed <= getNumValue)
                {
                    getIt = true;
                    simRatioDic[getNum] = simRatioDic[getNum] + 1;
                }
                else
                {
                    getNum++;
                }
            }
        }

        //输出到表格
        object[,] array2D = new object[simRatioDic.Count, 2];
        int row = 0;
        foreach (var kvp in simRatioDic)
        {
            array2D[row, 0] = kvp.Key; // 第一列放key
            array2D[row, 1] = kvp.Value; // 第二列放value
            row++;
        }

        PubMetToExcel.WriteExcelDataC(sheetName, 1, 1 + simRatioDic.Count - 1, 4, 5, array2D);
    }

    //向上获取最接近的key
    private static double GetNearKeyValue(Dictionary<int, double> dic, int getNum)
    {
        double getNumValue = 0;
        while (!dic.ContainsKey(getNum))
        {
            getNum--;
        }
        getNumValue = dic[getNum];
        return getNumValue;
    }

    //遍历目录文件
    public static void ExcelFolderPath(string[] folder)
    {
        var baseFolder = folder[1];
        var newPath = Path.GetDirectoryName(
            Path.GetDirectoryName(Path.GetDirectoryName(baseFolder))
        );
        if (newPath != null)
        {
            var filesCollection = new SelfExcelFileCollector(newPath);
            var baseFiles = filesCollection.GetAllExcelFilesPath();

            var sheetIndex = new List<string>();

            foreach (var baseFile in baseFiles)
            {
                var baseFileName = Path.GetFileName(baseFile);
                var basePath = Path.GetDirectoryName(baseFile);

                if (basePath != null && basePath.Contains("克朗代克"))
                {
                    baseFileName = "克朗代克##" + baseFileName;
                }

                //遍历Sheet
                var sheetNames = MiniExcel.GetSheetNames(baseFile);
                if (baseFileName.Contains("$"))
                {
                    //sheetNames = sheetNames.Where(name => !name.Contains("#")).ToList();
                    foreach (var name in sheetNames)
                    {
                        if (!name.Contains("#"))
                        {
                            sheetIndex.Add(baseFileName + "##" + name);
                        }
                    }
                }
                else
                {
                    sheetIndex.Add(baseFileName);
                }
            }

            var rowCount = sheetIndex.Count;
            // 创建一个二维数组，行数为list的长度，列数为1
            object[,] result = new object[sheetIndex.Count, 1];

            for (int i = 0; i < sheetIndex.Count; i++)
            {
                result[i, 0] = sheetIndex[i];
            }

            var workSheet = NumDesAddIn.App.ActiveSheet;
            if (workSheet.Name.Contains("文件目录"))
            {
                var targetRange = workSheet.Range["E6:E" + (rowCount + 5)];
                targetRange.Value = "";
                targetRange.Value = result;
            }
            else
            {
                MessageBox.Show("当前表格不是“#表格关联##文件目录，请切换");
            }
        }
    }

    public static void FormularBaseCheck()
    {
        var app = NumDesAddIn.App;
        var wk = app.ActiveWorkbook;
        var basePath = wk.Path;

        if (basePath.Contains("克朗代克"))
        {
            basePath = Path.GetDirectoryName(Path.GetDirectoryName(basePath));
        }

        basePath = Path.GetDirectoryName(basePath);
        var baseFilePathList = new List<string>();

        // 指定文件类型（扩展名）
        string[] fileTypes = { "*.xlsx", "*.xlsm" }; // 例如，获取 .txt, .csv 和 .xml 文件

        // 遍历每种文件类型
        foreach (string fileType in fileTypes)
        {
            // 获取指定目录及其子目录中的所有指定类型的文件
            if (basePath != null)
            {
                string[] files = Directory.GetFiles(
                    basePath,
                    fileType,
                    SearchOption.AllDirectories
                );

                // 将文件路径添加到集合中
                foreach (string file in files)
                {
                    baseFilePathList.Add(file);
                }
            }
        }

        var links = wk.LinkSources(XlLink.xlExcelLinks);
        if (links == null || links.Length == 0)
        {
            MessageBox.Show("没有检测到有外链公式");
            return;
        }

        var needFixLinks = new List<string>();
        foreach (string link in links)
        {
            if (!baseFilePathList.Contains(link))
            {
                var fileName = Path.GetFileName(link);
                var filePath = Path.GetDirectoryName(link);
                //根目录与非根目录统一格式
                if (!string.IsNullOrEmpty(filePath) && !filePath.EndsWith(@"\"))
                {
                    filePath += @"\";
                }

                var newLink = filePath + @"[" + fileName + @"]";
                needFixLinks.Add(newLink);
            }
        }

        if (needFixLinks.Count != 0)
        {
            var replaceFixLinks = new List<string>();
            InputFormularWindow inputDialog = new InputFormularWindow(needFixLinks);
            if (inputDialog.ShowDialog() == true)
            {
                replaceFixLinks = inputDialog.UserInputs;
            }

            if (replaceFixLinks != null)
            {
                NumDesAddIn.App.ScreenUpdating = false;
                NumDesAddIn.App.Calculation = XlCalculation.xlCalculationManual;

                // 遍历所有工作表
                foreach (Worksheet worksheet in wk.Worksheets)
                {
                    var wsName = worksheet.Name;
                    //跳过外链数据表
                    if (wsName.Contains("【外源数据】"))
                    {
                        continue;
                    }

                    // 遍历工作表中的所有单元格
                    Range usedRange = worksheet.UsedRange;
                    foreach (Range cell in usedRange)
                    {
                        if (cell.HasFormula)
                        {
                            // 获取原始公式
                            var originalFormula = cell.Formula;
                            var newFormula = originalFormula;
                            for (int indexFor = 0; indexFor < needFixLinks.Count; indexFor++)
                            {
                                var oldFor = needFixLinks[indexFor];
                                var newFor = replaceFixLinks[indexFor];

                                var checkNewFor = newFor.Replace("[", "").Replace("]", "");
                                if (baseFilePathList.Contains(checkNewFor))
                                {
                                    // 替换公式样式
                                    newFormula = newFormula.Replace(oldFor, newFor);

                                    // 不带[]替换
                                    var checkOldFor = oldFor.Replace("[", "").Replace("]", "");
                                    newFormula = newFormula.Replace(checkOldFor, checkNewFor);
                                }
                                else
                                {
                                    //修改的公式也是错的，跳过
                                    MessageBox.Show("修改的公式也是错误的，执行失败");
                                    return;
                                }
                            }

                            if (wsName.Contains("LTE【通用】"))
                            {
                                Debug.Print($"{wsName}:{cell.Address}");
                                Debug.Print($"原公式:{originalFormula}");
                                Debug.Print($"新公式:{newFormula}");
                            }

                            if (originalFormula != newFormula)
                            {
                                // 设置新的公式
                                cell.Formula = newFormula;
                            }
                        }
                    }

                    // 遍历嵌入在工作表中的图表
                    foreach (ChartObject chartObject in worksheet.ChartObjects())
                    {
                        var chart = chartObject.Chart;
                        foreach (Series series in chart.SeriesCollection())
                        {
                            string formula = series.Formula;
                            string newFormula = formula;
                            for (int indexFor = 0; indexFor < needFixLinks.Count; indexFor++)
                            {
                                var oldFor = needFixLinks[indexFor];
                                var newFor = replaceFixLinks[indexFor];
                                // 替换公式样式
                                newFormula = newFormula.Replace(oldFor, newFor);
                            }

                            if (formula != newFormula)
                            {
                                // 设置新的公式
                                series.Formula = newFormula;
                            }
                        }
                    }
                }

                NumDesAddIn.App.Calculation = XlCalculation.xlCalculationAutomatic;
                NumDesAddIn.App.ScreenUpdating = true;
            }
        }
    }

    public static void LoopRunCac(string sheetName)
    {
        var baseName = "转盘奔跑";
        var a1Row = 42;
        var a2Row = 66;
        var startRow = a1Row;
        if (sheetName == "转盘奔跑 【非大矿】")
        {
            startRow = a2Row;
        }

        if (!sheetName.Contains(baseName))
            MessageBox.Show("当前表格不是【转盘奔跑**】,无法使用转盘奔跑功能");
        var pointFunc = PubMetToExcel.ReadExcelDataC(sheetName, startRow - 13, startRow - 2, 2, 4);
        var checkInfo = PubMetToExcel.ReadExcelDataC(sheetName, startRow - 13, startRow - 9, 5, 5);
        var checkInfo2List = PubMetToExcel.Array2DDataToList(checkInfo);
        var checkInfoList = PubMetToExcel.List2DToListRowOrCol(checkInfo2List, true);

        LoopRunCheckBoxWindow checkWindow = new LoopRunCheckBoxWindow(checkInfoList);
        checkWindow.ShowDialog();
        var checkCurrent = checkWindow.SelectedList;
        if (checkCurrent == null)
        {
            return;
        }

        if (checkCurrent.Count != 0)
        {
            int multipleRank = 1;
            foreach (var unused in checkCurrent)
            {
                var averageRange = PubMetToExcel.ReadExcelDataC(
                    sheetName,
                    startRow + (multipleRank - 1) * 21,
                    startRow + (multipleRank - 1) * 21,
                    6,
                    7
                );
                var maxRollCell = PubMetToExcel.ReadExcelDataC(
                    sheetName,
                    startRow + (multipleRank - 1) * 21,
                    startRow + (multipleRank - 1) * 21,
                    9,
                    9
                );
                var average = Convert.ToDouble(pointFunc[0, 2]);
                var averageRangeMin = (Convert.ToDouble(averageRange[0, 0]) + 1) * average;
                var averageRangeMax = (Convert.ToDouble(averageRange[0, 1]) + 1) * average;
#pragma warning disable CA1305
                var maxRoll = Convert.ToInt32(maxRollCell[0, 0]);
#pragma warning restore CA1305

                var randFunList = PubMetToExcel.UniqueRandomMethod(maxRoll, maxRoll * 100000, 12);

                //二维数据字典化方便查找
                var pointFuncDic = PubMetToExcel.TwoDArrayToDictionary(pointFunc);

                Dictionary<int, List<int>> pointTotalList = new Dictionary<int, List<int>>();
                for (int i = 0; i < randFunList.Count; i++)
                {
                    var randFun = randFunList[i];
                    List<int> pointTotal = new List<int>();
                    var randMax = randFun.Count;
                    double pointCount = 0;
                    for (int j = 0; j < randMax; j++)
                    {
                        var randFunSeed = randFun[j];
                        pointCount += Convert.ToDouble(pointFuncDic[randFunSeed][1]);
                        pointTotal.Add(randFunSeed);
                    }

                    var pointAverage = pointCount / randMax;
                    if (
                        pointAverage >= averageRangeMin
                        && pointAverage <= averageRangeMax
                        && pointCount % 12 == 0
                    )
                    {
                        pointTotalList[i] = pointTotal;
                    }
                }

                //过滤至少包含某个值n次
                pointTotalList = pointTotalList
                    .Where(kvp => kvp.Value.Count(x => x == 1) >= 1)
                    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                pointTotalList = pointTotalList
                    .Where(kvp => kvp.Value.Count(x => x == 6) >= 1)
                    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                //随机获取字典中最多10个数据
                pointTotalList = PubMetToExcel.RandChooseDataFormDictionary(pointTotalList, 10);

                var pointArray = PubMetToExcel.DictionaryTo2DArray(
                    pointTotalList,
                    maxRows: 10,
                    maxCols: maxRoll
                );
                object[,] pointArrayStr = PubMetToExcel.ConvertToCommaSeparatedArray(pointArray);

                //清除老数据
                var emptyData = new object[10, 1];
                PubMetToExcel.WriteExcelDataC(
                    sheetName,
                    startRow + 2 + (multipleRank - 1) * 21,
                    startRow + 11 + (multipleRank - 1) * 21,
                    3,
                    3,
                    emptyData
                );
                //填写新数据
                PubMetToExcel.WriteExcelDataC(
                    sheetName,
                    startRow + 2 + (multipleRank - 1) * 21,
                    startRow + 11 + (multipleRank - 1) * 21,
                    3,
                    3,
                    pointArrayStr
                );
                multipleRank++;
            }
        }
    }

    //检查数据合法性
    public static void CheckDataLegitimacy(string rootPath)
    {
        //获取指定目录所有文件信息
        var filesCollector = new SelfExcelFileCollector(rootPath);
        var filesMd5 = filesCollector.GetAllExcelFilesMd5(
            SelfExcelFileCollector.KeyMode.FileNameWithoutExt
        );

        //读取Excel原始文件MD5值进行比对
        var md5FilePath =
            Path.GetDirectoryName(Path.GetDirectoryName(rootPath))
            + @"\Excels\ExcelRelationPath.txt";

        var changedFiles = new List<string>();
        foreach (var line in File.ReadLines(md5FilePath))
        {
            var parts = line.Split('|');
            if (parts.Length != 3)
                continue;

            var key = parts[0];
            var md5 = parts[1];

            if (filesMd5.TryGetValue(key, out var value) && value.MD5 != md5)
            {
                changedFiles.Add(value.FullPath);
            }
        }

        //获取所有变化表格数据
        //检查Md5不同文件的数据是否合法
        int currentCount = 0;
        int totalCount = changedFiles.Count;
        var errorList = new List<(string, string, string, SelfCellData)>();

        foreach (var changedFile in changedFiles)
        {
            var sheetNames = new List<string>();

            // 获取工作表名称
            sheetNames = MiniExcel.GetSheetNames(changedFile);

            //if (changedFile != @"C:\M1Work\Public\Excels\Tables\RewardGroup.xlsx")
            //{
            //    var abc = 0;
            //    continue;
            //}

            if (!changedFile.Contains("$"))
            {
                var realSheetName = sheetNames[0];
                if (sheetNames.Contains("Sheet1"))
                {
                    realSheetName = "Sheet1";
                }

                sheetNames = new List<string> { realSheetName };
            }

            foreach (var sheetName in sheetNames)
            {
                if (sheetName.Contains("#"))
                    continue;

                var rows = MiniExcel.Query(changedFile, sheetName: sheetName).ToList();

                var keyRow = rows[1] as IDictionary<string, object>;
                var typeRow = rows[2] as IDictionary<string, object>;
                var typeCols = new List<string>(typeRow.Keys);
                for (int rowIndex = 4; rowIndex < rows.Count; rowIndex++)
                {
                    var row = rows[rowIndex] as IDictionary<string, object>;

                    for (int colIndex = 2; colIndex < typeCols.Count; colIndex++)
                    {
                        var col = typeCols[colIndex - 1];
                        var typeCell = typeRow[col]?.ToString() ?? "";
                        var keyCell = keyRow[col]?.ToString() ?? "";
                        if (keyCell == "")
                        {
                            continue;
                        }

                        if (typeCell != null)
                        {
                            if (typeCell.ToString().Contains("#"))
                            {
                                continue;
                            }

                            var typeData = typeCell.ToString().Split('=');
                            var typeName = typeData[0]?.ToString() ?? "";
                            var defaultValue = "";
                            var cell = row[col]?.ToString() ?? "";
                            if (typeData.Length > 1)
                            {
                                defaultValue = typeData[1];
                            }

                            var typeCellData = new SelfCellData((typeCell.ToString(), 2, colIndex));
                            var cellData = new SelfCellData((cell, rowIndex + 1, colIndex));

                            //object[]就是：number[](double[])和string[]；如果值里面没有[],默认加上，[]无脑换成{}；
                            //object[][]和table都是table
                            //判断类型是否不是一维数组
                            if (
                                typeName.Contains("[]")
                                && !typeName.Contains("table")
                                && !typeName.Contains("[][]")
                            )
                            {
                                if (cell == "" || cell == "[]" || cell == "[][]")
                                    continue;
                                //没有[]的单值需要补充[]
                                if (!cell.Contains("["))
                                    cell = "[" + cell + "]";
                                //检查并解析数组，Value值类型没法检测，因为有人混用数字、字符
                                if (!PubMetToExcel.IsValidArray(cell, out object[] array))
                                {
                                    var errorTips = $"数组格式错误：{typeName}";
                                    errorList.Add((changedFile, sheetName, errorTips, cellData));
                                }
                            }
                            //判断类型是否不是table和二维数组
                            else if (typeName.Contains("table") || typeName.Contains("[][]"))
                            {
                                if (cell == "" || cell == "{}")
                                    continue;

                                //二维转table
                                if (typeName.Contains("[][]"))
                                {
                                    //object[][]都是table
                                    cell = cell.Replace("[", "{").Replace("]", "}");
                                }

                                //没有{}的单值需要补充{}
                                if (!cell.Contains("{"))
                                    cell = "{" + cell + "}";

                                // 检查并解析LuaTable
                                using (Lua lua = new Lua())
                                {
                                    // 检查 value
                                    try
                                    {
                                        lua.DoString($"value = {cell}");
                                        if (!(lua["value"] is LuaTable))
                                        {
                                            var errorTips = $"数组/表格格式错误：{typeName}";
                                            errorList.Add(
                                                (changedFile, sheetName, errorTips, cellData)
                                            );
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        var errorTips = $"数组/表格格式错误：{typeName}#【{ex.Message}】";
                                        errorList.Add(
                                            (changedFile, sheetName, errorTips, cellData)
                                        );
                                    }
                                }
                            }
                            //其他一般类型，Value值类型没法检测，因为有人混用数字、字符
                        }
                    }
                }
            }

            currentCount++;
            NumDesAddIn.App.StatusBar = $"正在检查第 {currentCount}/{totalCount} 个文件: {changedFile}";
        }
    }

    //更新Power Query链接数据
    public static void UpdatePowerQueryLinks()
    {
        dynamic queries = null;
        try
        {
            // 新的文件夹路径
            var newFolderPath = WkPath + @"\";

            // 获取工作簿中的所有查询
            queries = Wk.GetType()
                .InvokeMember("Queries", BindingFlags.GetProperty, null, Wk, null);

            foreach (dynamic query in queries)
            {
                string name = query.Name;
                string formula = query.Formula;

                // 提取旧文件路径
                string oldFilePath = ExtractFilePathFromFormula(formula);
                if (!string.IsNullOrEmpty(oldFilePath))
                {
                    // 获取文件名并生成新的文件路径
                    string fileName = Path.GetFileName(oldFilePath);
                    string newFilePath = Path.Combine(newFolderPath, fileName);

                    // 替换旧路径为新路径
                    string newFormula = formula.Replace(oldFilePath, newFilePath);

                    // 更新查询公式
                    query
                        .GetType()
                        .InvokeMember(
                            "Formula",
                            BindingFlags.SetProperty,
                            null,
                            query,
                            new object[] { newFormula }
                        );
                }
            }

            // 刷新所有查询
            Wk.GetType().InvokeMember("RefreshAll", BindingFlags.InvokeMethod, null, Wk, null);
        }
        catch (Exception ex)
        {
            // 捕获并处理异常
            MessageBox.Show("更新 Power Query 链接时发生错误: " + ex.Message);
        }
        finally
        {
            // 释放查询对象，避免锁定外部文件
            if (queries != null)
            {
                Marshal.ReleaseComObject(queries);
                queries = null;
            }

            // 强制垃圾回收，确保释放未使用的 COM 对象
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    //更新Power Query链接数据-提取公式
    private static string ExtractFilePathFromFormula(string formula)
    {
        // 提取文件路径的逻辑
        // 假设文件路径在公式中以 "File.Contents(" 开头，并以 ")" 结尾
        string startMarker = "File.Contents(\"";
        string endMarker = "\")";
        int startIndex = formula.IndexOf(startMarker) + startMarker.Length;
        int endIndex = formula.IndexOf(endMarker, startIndex);
        if (startIndex > startMarker.Length && endIndex > startIndex)
        {
            return formula.Substring(startIndex, endIndex - startIndex);
        }

        return string.Empty;
    }

    //替换查找的字符
    public static void ReplaceValueFormat(string specialCharsStr)
    {
        var indexWk = NumDesAddIn.App.ActiveWorkbook;

        var sourceSheet = indexWk.Worksheets["Sheet1"];

        var sourceMaxCol = sourceSheet.UsedRange.Columns.Count;
        var sourceMaxRow = sourceSheet.UsedRange.Rows.Count;
        var sourceRange = sourceSheet.Range[
            sourceSheet.Cells[5, 7],
            sourceSheet.Cells[sourceMaxRow, sourceMaxCol]
        ];
        var sourceData = new List<(string, int, int)>();
        // 将 specialCharsStr 中的转义字符转换为实际字符
        var specialChars = specialCharsStr.Split('#').Select(Regex.Unescape).ToArray();
        for (int col = 1; col <= sourceMaxCol - 7 + 1; col++)
        {
            for (int row = 1; row <= sourceMaxRow - 5 + 1; row++)
            {
                var cell = sourceRange[row, col];

                // 检查单元格的字符串中是否有换行符
                string cellValue = cell.Value2?.ToString() ?? "";
                bool hasNewChar = specialChars.Any(cellValue.Contains);
                if (!hasNewChar)
                    continue;
                int cellRow = cell.Row;
                int cellCol = cell.Column;

                sourceData.Add((cellValue, cellRow, cellCol));

                //替换字符串
                var replaceedValue = "";
                for (int i = 0; i < specialChars.Length; i++)
                {
                    cellValue = cellValue.Replace(specialChars[i], replaceedValue);
                }

                cell.Value2 = cellValue;
            }
        }

        if (sourceRange.Count == 0)
        {
            MessageBox.Show("未找到匹配值");
        }
        else
        {
            var ctpName = "表格查询结果";
            NumDesCTP.DeleteCTP(true, ctpName);
            _ = (CellSeachResult)
                NumDesCTP.ShowCTP(
                    550,
                    ctpName,
                    true,
                    ctpName,
                    new CellSeachResult(sourceData),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    //查找字符
    public static void SeachValueFormat(string specialCharsStr)
    {
        var indexWk = NumDesAddIn.App.ActiveWorkbook;

        var sourceSheet = indexWk.Worksheets["Sheet1"];

        var sourceMaxCol = sourceSheet.UsedRange.Columns.Count;
        var sourceMaxRow = sourceSheet.UsedRange.Rows.Count;
        var sourceRange = sourceSheet.Range[
            sourceSheet.Cells[5, 7],
            sourceSheet.Cells[sourceMaxRow, sourceMaxCol]
        ];
        var sourceData = new List<(string, int, int)>();
        // 将 specialCharsStr 中的转义字符转换为实际字符
        var specialChars = specialCharsStr.Split('#').Select(Regex.Unescape).ToArray();
        for (int col = 1; col <= sourceMaxCol - 7 + 1; col++)
        {
            for (int row = 1; row <= sourceMaxRow - 5 + 1; row++)
            {
                var cell = sourceRange[row, col];

                // 检查单元格的字符串中是否有换行符
                string cellValue = cell.Value2?.ToString() ?? "";
                bool hasNewChar = specialChars.Any(cellValue.Contains);
                if (!hasNewChar)
                    continue;
                int cellRow = cell.Row;
                int cellCol = cell.Column;
                sourceData.Add((cellValue, cellRow, cellCol));
            }
        }

        if (sourceRange.Count == 0)
        {
            MessageBox.Show("未找到匹配值");
        }
        else
        {
            var ctpName = "表格查询结果";
            NumDesCTP.DeleteCTP(true, ctpName);
            _ = (CellSeachResult)
                NumDesCTP.ShowCTP(
                    550,
                    ctpName,
                    true,
                    ctpName,
                    new CellSeachResult(sourceData),
                    MsoCTPDockPosition.msoCTPDockPositionRight
                );
        }
    }

    //数据查重-MiniExcel
    public static List<(string, int, int, string, string)> CheckRepeatValue(
        IEnumerable<dynamic> rows,
        string sheetName
    )
    {
        var sourceData = new List<(string, int, int, string, string)>();

        var dataRows = rows.Skip(3).ToList();

        if (dataRows.Count == 0)
        {
            return sourceData;
        }

        // 检查第 1、2 列第 1 行的值是否为特定字符串，如果是则跳过该工作表
        if (
            dataRows.Any() && ((IDictionary<string, object>)dataRows[0])["A"]?.ToString() != "#"
            || ((IDictionary<string, object>)dataRows[0])["B"]?.ToString() == null
        )
        {
            return sourceData;
        }

        // 检查 List 中第 2 列是否有重复值，并返回重复值的行列号
        var duplicates = dataRows
            .AsParallel()
            .Select((row, index) => new { Row = row, Index = index + 4 }) // 保留行号，+5 是因为跳过了前 4 行
            .Where(x => ((IDictionary<string, object>)x.Row)["B"] != null) // 忽略 null 值
            .GroupBy(x => ((IDictionary<string, object>)x.Row)["B"]) // 按第 2 列的值分组
            .Where(group => group.Count() > 1) // 找出重复值
            .SelectMany(group => group) // 展开分组
            .ToList();

        //转换数据格式
        foreach (var duplicate in duplicates)
        {
            var cellValue = ((IDictionary<string, object>)duplicate.Row)["B"].ToString();
            var cellRow = duplicate.Index;
            var cellCol = 2; // 第 2 列
            sourceData.Add((cellValue, cellRow, cellCol, sheetName, "数据重复"));
        }

        return sourceData;
    }

    //数据格式检查-MiniExcel
    public static List<(string, int, int, string, string)> CheckValueFormat(
        IEnumerable<dynamic> rows,
        string sheetName
    )
    {
        var sourceData = new List<(string, int, int, string, string)>();

        //有可能不需要这么复杂的判断，只判断是否包含常见的错误组合
        //比如【双逗号，中括号+逗号，大括号+逗号】
        //数组判断就通过头字符是否是{{、{、[[、[来检查末尾是否有对应的符号

        //var charactersToCheck = new[]
        //{
        //    ",,",
        //    "[,",
        //    ",]",
        //    "{,",
        //    ",}",
        //    "，，",
        //    "[，",
        //    "，]",
        //    "{，",
        //    "，}",
        //    "][",
        //    "}{"
        //};

        //var stringPairs = new List<(string leftString, string rightString)>
        //{
        //    ("[", "]"),
        //    ("{", "}")
        //};

        var config = new GlobalVariable();
        var normalCharactersCheck = config.NormaKeyList;
        var specialCharactersCheck = config.SpecialKeyList;
        var coupleCharactersCheck = config.CoupleKeyList;

        var dataRows = rows.ToList();

        if (dataRows.Count == 0)
        {
            return sourceData;
        }

        // 检查第 1、2 列第 1 行的值是否为特定字符串，如果是则跳过该工作表
        if (
            dataRows.Any() && ((IDictionary<string, object>)dataRows[3])["A"]?.ToString() != "#"
            || ((IDictionary<string, object>)dataRows[3])["B"]?.ToString() == null
        )
        {
            return sourceData;
        }

        var keyRow = dataRows[1] as IDictionary<string, object>;
        var keyType = dataRows[2] as IDictionary<string, object>;
        var keyCols = new List<string>(keyRow.Keys);
        for (int rowIndex = 4; rowIndex < dataRows.Count; rowIndex++)
        {
            var row = dataRows[rowIndex] as IDictionary<string, object>;

            for (int colIndex = 2; colIndex < keyCols.Count; colIndex++)
            {
                var col = keyCols[colIndex];
                var keyCell = keyRow[col]?.ToString() ?? "";
                var typeCell = keyType[col]?.ToString() ?? "";
                if (keyCell == "" || keyCell.Contains("#"))
                {
                    continue;
                }

                var cellValue = row[col]?.ToString();
                if (cellValue != null)
                {
                    // int数据判断
                    var types = typeCell.Split('=');
                    if (types[0] == "int" || types[0] == "int[]")
                    {
                        var cellValueSplit = cellValue
                            .Split(',', StringSplitOptions.RemoveEmptyEntries)
                            .Select(s =>
                            {
                                var match = Regex.Match(s, @"\d+(?:\.\d+)?");
                                return match.Success ? match.Value : null;
                            })
                            .Where(s => !string.IsNullOrEmpty(s))
                            .ToList();

                        foreach (var cellSplit in cellValueSplit)
                        {
                            // 尝试解析为浮点数，判断是否为整数
                            if (double.TryParse(cellSplit, out double cellSplitDouble))
                            {
                                // 检查是否为整数（浮点数的小数部分为0）
                                if (cellSplitDouble % 1 != 0)
                                {
                                    sourceData.Add(
                                        (
                                            cellValue,
                                            rowIndex + 1,
                                            colIndex + 1,
                                            sheetName,
                                            $"【{typeCell}】格式错误"
                                        )
                                    );
                                    break;
                                }
                            }
                        }
                    }
                    // 其他非法字符
                    if (
                        normalCharactersCheck.Any(c => cellValue.Contains(c))
                        && !typeCell.Contains("string")
                    )
                    {
                        sourceData.Add(
                            (cellValue, rowIndex + 1, colIndex + 1, sheetName, "多逗号或中文逗号")
                        );
                    }

                    if (
                        specialCharactersCheck.Any(c => cellValue.Contains(c))
                        && !typeCell.Contains("string")
                    )
                    {
                        sourceData.Add((cellValue, rowIndex + 1, colIndex + 1, sheetName, "少逗号"));
                    }

                    foreach (var (leftString, rightString) in coupleCharactersCheck)
                    {
                        var leftStringCount = Regex
                            .Matches(cellValue, Regex.Escape(leftString), RegexOptions.IgnoreCase)
                            .Count;
                        var RightStringCount = Regex
                            .Matches(cellValue, Regex.Escape(rightString), RegexOptions.IgnoreCase)
                            .Count;
                        if (leftStringCount != RightStringCount)
                        {
                            sourceData.Add(
                                (cellValue, rowIndex + 1, colIndex + 1, sheetName, "括号问题")
                            );
                            break;
                        }

                        if (leftString == "\"")
                        {
                            int isDouble = leftStringCount % 2;
                            if (isDouble != 0)
                            {
                                sourceData.Add(
                                    (cellValue, rowIndex + 1, colIndex + 1, sheetName, "双引号问题")
                                );
                                break;
                            }
                        }
                    }
                }
            }
        }

        return sourceData;
    }

    //数组数据格式检查-MiniExcel
    public static string CheckArrayValueFormat(
        string sheetName,
        string checkCol,
        string wkFullPath,
        string targetWkName,
        string targetSheetName,
        string checkTargetCol,
        string errorTips
    )
    {
        string result = string.Empty;

        var rows = MiniExcel
            .Query(
                wkFullPath,
                sheetName: sheetName,
                configuration: NumDesAddIn.OnOffMiniExcelCatches,
                startCell: "A2",
                useHeaderRow: true
            )
            .ToList();

        if (rows.Count == 0)
            return null;

        var targetPath = Path.Combine(Path.GetDirectoryName(wkFullPath), targetWkName);
        var rowsTarget = MiniExcel
            .Query(
                targetPath,
                sheetName: targetSheetName,
                configuration: NumDesAddIn.OnOffMiniExcelCatches,
                startCell: "A2",
                useHeaderRow: true
            )
            .ToList();
        if (rowsTarget.Count == 0)
            return null;

        // 遍历指定字段的数组数据
        for (int rowIndex = 2; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex] as IDictionary<string, object>;

            var cellIndex = row["id"]?.ToString();
            var cellComment = row["#备注"]?.ToString();
            var cellValue = row[checkCol]?.ToString();

            if (cellValue == null)
                continue;

            cellValue = cellValue.Replace("[", "");
            cellValue = cellValue.Replace("]", "");

            if (cellValue == "")
                continue;

            var cellValueGroup = cellValue.Split(",");

            var resultList = new List<string>();

            foreach (var checkId in cellValueGroup)
            {
                // 目标数据中检查是否合法
                for (int i = 2; i < rowsTarget.Count; i++)
                {
                    var rowTarget = rowsTarget[i] as IDictionary<string, object>;
                    var cellTargetIndex = rowTarget["id"]?.ToString();
                    if (checkId == cellTargetIndex)
                    {
                        var cellTargetValue = rowTarget[checkTargetCol]?.ToString();
                        if (!string.IsNullOrEmpty(cellTargetValue))
                        {
                            resultList.Add(cellTargetIndex);
                        }
                        break;
                    }
                }
            }

            if (resultList.Count > 0)
            {
                result +=
                    $"id:{cellIndex}#:{cellComment}# {errorTips}:{string.Join(",", resultList)}\n";
            }
        }

        return result;
    }

    //Excel名称表数据写入
    public static void UpdateExcelNameData(
        Worksheet sheet,
        string listObjectName,
        List<object> newData
    )
    {
        var listObject = sheet.ListObjects[listObjectName];
        var dataRange = listObject.DataBodyRange;
        int currentRowCount = dataRange.Rows.Count;
        int newRowCount = newData.Count;

        // 调整行数
        if (newRowCount > currentRowCount)
        {
            // 添加行
            for (int i = currentRowCount + 1; i <= newRowCount; i++)
            {
                listObject.ListRows.Add(Type.Missing);
            }
        }
        else if (newRowCount < currentRowCount)
        {
            // 删除多余行
            for (int i = currentRowCount; i > newRowCount; i--)
            {
                listObject.ListRows[i].Delete();
            }
        }

        // 写入新数据到第1列，并复制其他列数据
        for (int i = 0; i < newRowCount; i++)
        {
            dataRange.Cells[i + 1, 1].Value2 = newData[i];
            for (int j = 2; j <= dataRange.Columns.Count; j++)
            {
                dataRange.Cells[i + 1, j].Value2 = dataRange.Cells[i + 1, j].Value2;
            }
        }
    }

    //砸冰块计算
    public static void IceClimberCostSimulate(string wkPath, dynamic wk)
    {
        var modelSheetName = "#冰块模版";
        var posSheetName = "IceClimberTargetCell";
        var sizeSheetName = "IceClimberTargetTemp";

        // 查询模版数据
        var modelRows = MiniExcel.Query(
            wkPath,
            sheetName: modelSheetName,
            startCell: "A1",
            useHeaderRow: true
        );

        // 查询位置数据
        var posRows = MiniExcel.Query(
            wkPath,
            sheetName: posSheetName,
            startCell: "A2",
            useHeaderRow: true
        );

        // 查询尺寸数据
        var sizeRows = MiniExcel.Query(
            wkPath,
            sheetName: sizeSheetName,
            startCell: "A2",
            useHeaderRow: true
        );

        // 存储结果
        var modelResults = new Dictionary<string, List<(string, string, string, string)>>();

        foreach (var item in modelRows)
        {
            var column1 = item.模版名.ToString();
            var column2Values = new List<string>(ExtractNumbers(item.模版目标.ToString()));
            var column3Values = new List<string>(ExtractNumbers(item.模版尺寸.ToString()));

            var resultList = new List<(string, string, string, string)>();

            int maxCount = Math.Max(column2Values.Count, column3Values.Count);

            for (int i = 0; i < maxCount; i++)
            {
                var posA = "";
                var posB = "";
                var sizeA = "";
                var sizeB = "";

                if (i < column2Values.Count)
                {
                    var matchingPosRow = posRows.FirstOrDefault(row =>
                        row.id.ToString() == column2Values[i]
                    );
                    if (matchingPosRow != null)
                    {
                        posA = matchingPosRow.start_x.ToString();
                        posB = matchingPosRow.start_y.ToString();
                    }
                }

                if (i < column3Values.Count)
                {
                    var matchingSizeRow = sizeRows.FirstOrDefault(row =>
                        row.id.ToString() == column3Values[i]
                    );
                    if (matchingSizeRow != null)
                    {
                        sizeB = matchingSizeRow.wide.ToString();
                        sizeA = matchingSizeRow.high.ToString();
                    }
                }

                resultList.Add((posA, posB, sizeA, sizeB));
            }

            if (resultList.Any())
            {
                modelResults[column1] = resultList;
            }
        }

        int totalRows = 0;
        int totalCols = 0;
        foreach (var model in modelResults)
        {
            var modelValues = model.Value;
            int modelSize = modelValues.Count;
            totalRows += modelSize + 1;
            totalCols = Math.Max(totalCols, modelSize);
        }

        var combinedGrid = new string[totalRows, totalCols];

        int currentRow = 0;
        foreach (var model in modelResults)
        {
            var modelKey = model.Key;
            var modelValues = model.Value;
            int modelSize = modelValues.Count;

            var grid = new string[modelSize + 1, modelSize];

            grid[modelSize, 0] = modelKey;

            foreach (var modelValue in modelValues)
            {
                int startX = int.Parse(modelValue.Item1) - 1;
                int startY = int.Parse(modelValue.Item2) - 1;
                int width = int.Parse(modelValue.Item3);
                int height = int.Parse(modelValue.Item4);

                for (int i = startX; i < startX + width; i++)
                {
                    for (int j = startY; j < startY + height; j++)
                    {
                        grid[i, j] = "1";
                    }
                }
            }

            for (int i = 0; i < modelSize + 1; i++)
            {
                for (int j = 0; j < modelSize; j++)
                {
                    combinedGrid[currentRow + i, j] = grid[i, j];
                }
            }

            currentRow += modelSize + 1;
        }

        var outSheetName = "#冰块图形";
        var outSheet = wk.Sheets[outSheetName];
        var outRange = outSheet.Range[
            outSheet.Cells[2, 2],
            outSheet.Cells[1 + combinedGrid.GetLength(0), 1 + combinedGrid.GetLength(1)]
        ];
        outRange.Value = combinedGrid;
    }

    static IEnumerable<string> ExtractNumbers(string input)
    {
        // 使用正则表达式提取数字
        var matches = Regex.Matches(input, @"\d+");
        return matches.Cast<Match>().Select(m => m.Value);
    }

    // 同步Icon修正数据
    public static void SyncIconFixData(string filePath)
    {
        // 检查路径文件是否打开
        try
        {
            // 尝试以读写方式打开文件
            using (
                FileStream fs = File.Open(
                    filePath,
                    FileMode.Open,
                    FileAccess.ReadWrite,
                    FileShare.None
                )
            )
                ;
        }
        catch (IOException)
        {
            MessageBox.Show($"{filePath}已打开，请关闭");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"检查文件状态时出错: {ex.Message}");
        }

        // 获取剪切板上的数据
        var iconFixData = new Dictionary<string, string>();

        if (!Clipboard.ContainsText())
        {
            Debug.Print("剪切板中没有文本数据");
            MessageBox.Show("剪切板中没有文本数据");
        }

        string clipboardText = Clipboard.GetText();
        string[] lines = clipboardText.Split(
            new[] { '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries
        );

        if (lines.Length == 0)
        {
            Debug.Print("剪切板中没有有效数据");
            MessageBox.Show("剪切板中没有有效数据");
        }

        foreach (string line in lines)
        {
            // 跳过Unity调试日志行
            if (
                line.Contains("UnityEngine.Debug:Log")
                || line.Contains("UIIconImageInEditorScene")
                || line.Contains("ObjectBuilderEditor")
                || line.Contains("GUIUtility:ProcessEvent")
            )
            {
                continue;
            }

            string[] parts = line.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 2)
            {
                string id = parts[0];
                string value = parts[1];
                iconFixData[id] = value;
            }
        }

        if (iconFixData.Count == 0)
        {
            Debug.Print("剪切板中没有有效数据");
            MessageBox.Show("剪切板中没有有效数据");
            return;
        }
        // 按照key列匹配数据写入，根据Key的前缀，只针对当期活动的数据进行写入
        string fileName = Path.GetFileName(filePath);
        filePath = Path.GetDirectoryName(filePath);

        PubMetToExcel.SetExcelObjectEpPlus(
            filePath,
            fileName,
            out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel
        );

        int firstMatchRow = 5;
        // 最大行
        int endRow = targetSheet.Dimension.End.Row;

        //// 活动编号
        //string commonPrefix = string.Empty;
        //commonPrefix = iconFixData.Keys.Last().Substring(0,5);

        //for (int row = 1; row <= endRow; row++)
        //{
        //    string cellValue = targetSheet.Cells[row, 2].Text;

        //    // 检查前N个字符是否匹配（N = searchPrefix.Length）
        //    if (cellValue.StartsWith(commonPrefix))
        //    {
        //        if (firstMatchRow == -1)
        //            firstMatchRow = row; // 记录第一个匹配行

        //        lastMatchRow = row; // 更新最后一个匹配行
        //    }
        //}

        // 获取字段列
        var fieldKey = "spriteName";
        var fieldValue1 = "scale_offset_handbook";
        var fieldValue2 = "scale_offset_same";

        int fieldKeyCol = -1;
        int fieldValue1Col = -1;
        int fieldValue2Col = -1;

        // 最大列
        int endCol = targetSheet.Dimension.End.Column;
        for (int col = 1; col <= endCol; col++)
        {
            string headerValue = targetSheet.Cells[2, col].Text;
            if (headerValue == fieldKey)
            {
                fieldKeyCol = col;
            }
            else if (headerValue == fieldValue1)
            {
                fieldValue2Col = col;
            }
            else if (headerValue == fieldValue2)
            {
                fieldValue1Col = col;
            }
        }

        //if (firstMatchRow == -1 || lastMatchRow == -1)
        //{
        //    Debug.Print("未找到匹配的行");
        //    MessageBox.Show("未找到匹配的行");
        //    return;
        //}

        if (fieldKeyCol == -1 || fieldValue1Col == -1 || fieldValue2Col == -1)
        {
            Debug.Print("未找到匹配的列");
            MessageBox.Show("未找到匹配的列");
            return;
        }

        bool isWrite = false;

        for (int row = firstMatchRow; row <= endRow; row++)
        {
            var iconKey = targetSheet.Cells[row, fieldKeyCol].Text;
            if (iconFixData.ContainsKey(iconKey))
            {
                var iconValue1 = iconFixData[iconKey];
                targetSheet.Cells[row, fieldValue1Col].Value = iconValue1;
                targetSheet.Cells[row, fieldValue2Col].Value = iconValue1;
                isWrite = true;
            }
        }

        if (isWrite)
        {
            MessageBox.Show($"已全量匹配所有目标Id的图片，提交时注意分辨是否为自己主观更改！！！");

            targetExcel.Save();
        }
        else
        {
            MessageBox.Show($"没找到匹配的图片Id");
        }
    }

    #region Excel数据查找

    //Epplus搜索
    public static List<(string, string, int, int)> SearchKeyFromExcel(
        string rootPath,
        string findValue
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new List<(string, string, int, int)>();
        var currentCount = 0;
        var count = files.Length;
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        foreach (var file in files)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    try
                    {
                        var wk = package.Workbook;
                        for (var sheetIndex = 0; sheetIndex < wk.Worksheets.Count; sheetIndex++)
                        {
                            var sheet = wk.Worksheets[sheetIndex];
                            if (
                                sheet.Name.Contains("#")
                                || sheet.Name.Contains("Sheet") && sheet.Name != "Sheet1"
                            )
                                continue;
                            int rowMax = Math.Max(sheet.Dimension.End.Row, 4);
                            int colMax = Math.Max(sheet.Dimension.End.Column, 2);
                            for (var col = 2; col <= colMax; col++)
                            for (var row = 4; row <= rowMax; row++)
                            {
                                var cellValue = sheet.Cells[row, col].Value?.ToString();
                                var cellAddress = new ExcelCellAddress(row, col);
                                var cellCol = cellAddress.Column;
                                var cellRow = cellAddress.Row;

                                if (
                                    cellValue != null
                                    && (
                                        isAll
                                            ? cellValue.Contains(findValue)
                                            : cellValue == findValue
                                    )
                                )
                                {
                                    targetList.Add((file, sheet.Name, cellRow, cellCol));
                                }
                            }
                        }
                    }
                    catch
                    {
                        // 记录异常信息，继续处理下一个文件
                    }
                }
            }
            catch
            {
                // 记录异常信息，继续处理下一个文件
            }

            currentCount++;
            NumDesAddIn.App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
        }

        return targetList;
    }

    public static List<(string, string, int, int)> SearchKeyFromExcelMulti(
        string rootPath,
        string findValue
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new List<(string, string, int, int)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Parallel.ForEach(
            files,
            options,
            file =>
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(file)))
                    {
                        try
                        {
                            var wk = package.Workbook;
                            for (var sheetIndex = 0; sheetIndex < wk.Worksheets.Count; sheetIndex++)
                            {
                                var sheet = wk.Worksheets[sheetIndex];
                                if (
                                    sheet.Name.Contains("#")
                                    || sheet.Name.Contains("Sheet") && sheet.Name != "Sheet1"
                                )
                                    continue;
                                int rowMax = Math.Max(sheet.Dimension.End.Row, 4);
                                int colMax = Math.Max(sheet.Dimension.End.Column, 2);
                                for (var col = 2; col <= colMax; col++)
                                for (var row = 4; row <= rowMax; row++)
                                {
                                    var cellValue = sheet.Cells[row, col].Value?.ToString();
                                    var cellAddress = new ExcelCellAddress(row, col);
                                    var cellCol = cellAddress.Column;
                                    var cellRow = cellAddress.Row;

                                    if (
                                        cellValue != null
                                        && (
                                            isAll
                                                ? cellValue.Contains(findValue)
                                                : cellValue == findValue
                                        )
                                    )
                                    {
                                        targetList.Add((file, sheet.Name, cellRow, cellCol));
                                    }
                                }
                            }
                        }
                        catch
                        {
                            // 记录异常信息，继续处理下一个文件
                        }
                    }
                }
                catch
                {
                    // 记录异常信息，继续处理下一个文件
                }
            }
        );
        return targetList;
    }

    //Epplus检查是否包含多余列
    public static List<(string, string, int, int)> CheckColFromExcelMulti(
        string rootPath,
        bool isMulti = true
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new List<(string, string, int, int)>();

        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Action<string> processFile = file =>
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    if (file.Contains("#"))
                    {
                        return;
                    }

                    Debug.Print($"检查：{file}");
                    //if (file.Contains("$多人建造活动.xlsx"))
                    //{
                    //    var ac = 1;
                    //}
                    try
                    {
                        var wk = package.Workbook;

                        // 获取当前活动工作表
                        var activeSheetIndex = wk.View.ActiveTab;
                        var activeSheet = wk.Worksheets[activeSheetIndex + 1];

                        var isWrite = false;

                        for (var sheetIndex = 0; sheetIndex < wk.Worksheets.Count; sheetIndex++)
                        {
                            var sheet = wk.Worksheets[sheetIndex];
                            if (sheet.Name.Contains("#"))
                                continue;
                            if (sheet.Name == "Chart1")
                            {
                                //删除多余列
                                targetList.Add((file, sheet.Name + "：非法表", 2, 1));
                                wk.Worksheets.Delete(sheet.Name);
                                isWrite = true;

                                continue;
                            }
                            if (sheet.Name.Contains("Sheet") && file.Contains("$"))
                                continue;
                            if (sheet.Cells[2, 1].Value?.ToString() != "#")
                                continue;
                            int colMax = Math.Max(sheet.Dimension.End.Column, 2);

                            for (var col = colMax; col >= 1; col--)
                            {
                                // 空列检测
                                var cellValue = sheet.Cells[2, col].Value;
                                if (
                                    ReferenceEquals(cellValue, "")
                                    || ReferenceEquals(cellValue, " ")
                                )
                                {
                                    Debug.Print($"{file}:[{sheet.Name}]冗余列{col}/{colMax}");
                                    //删除多余列
                                    targetList.Add((file, sheet.Name + "：冗余列", 2, col));

                                    sheet.DeleteColumn(col);

                                    isWrite = true;
                                }

                                // 隐藏列检测
                                var colObj = sheet.Column(col);
                                if (colObj.Hidden)
                                {
                                    Debug.Print($"{file}:[{sheet.Name}]隐藏列{col}/{colMax}");

                                    targetList.Add((file, sheet.Name + "：隐藏列", 2, col));

                                    colObj.Hidden = false;

                                    isWrite = true;
                                }
                            }

                            // 整理格式
                            if (isWrite)
                            {
                                var range = sheet.Cells[sheet.Dimension.Address];
                                // 设置字体格式
                                range.Style.Font.Name = "微软雅黑";
                                range.Style.Font.Size = 10;

                                // 设置对齐方式
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                targetList.Add((file, sheet.Name + "：整理格式", 2, 2));

                                // 检测是否有多个Sheet被选中
                                sheet.View.TabSelected = false;
                            }
                        }
                        if (isWrite)
                        {
                            activeSheet.View.TabSelected = true;
                            package.Save();
                        }
                    }
                    catch
                    {
                        // 记录异常信息，继续处理下一个文件
                    }
                }
            }
            catch
            {
                // 记录异常信息，继续处理下一个文件
            }
        };

        if (isMulti)
        {
            Parallel.ForEach(files, options, processFile);
        }
        else
        {
            foreach (var file in files)
            {
                processFile(file);
            }
        }
        return targetList;
    }

    //MiniExcel查询：全局查询
    public static List<(string, string, int, int)> SearchKeyFromExcel(
        string rootPath,
        string findValue,
        bool isMulti = false,
        bool searchSpecificColumn = false,
        int specificColumnIndex = 2
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new ConcurrentBag<(string, string, int, int)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");

        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Action<string> processFile = file =>
        {
            try
            {
                var sheetNames = MiniExcel.GetSheetNames(file);
                foreach (var sheetName in sheetNames)
                {
                    if (sheetName.Contains("#"))
                        continue;

                    var rows = MiniExcel.Query(
                        file,
                        sheetName: sheetName,
                        configuration: NumDesAddIn.OnOffMiniExcelCatches
                    );
                    int rowIndex = 1;
                    foreach (var row in rows)
                    {
                        int colIndex = 1;
                        foreach (var cell in row)
                        {
                            if (searchSpecificColumn && colIndex != specificColumnIndex)
                            {
                                colIndex++;
                                continue;
                            }

                            var cellValue = cell.Value?.ToString();
                            if (
                                cellValue != null
                                && (isAll ? cellValue.Contains(findValue) : cellValue == findValue)
                            )
                            {
                                targetList.Add((file, sheetName, rowIndex, colIndex));
                            }

                            colIndex++;
                        }

                        rowIndex++;
                    }
                }
            }
            catch
            {
                // 记录异常信息，继续处理下一个文件
            }
        };

        if (isMulti)
        {
            Parallel.ForEach(files, options, processFile);
        }
        else
        {
            foreach (var file in files)
            {
                processFile(file);
            }
        }

        return targetList.ToList();
    }

    //MiniExcel查询：同一个表查找多个值
    public static Dictionary<string, List<string>> SearchKeysFrom1ExcelMulti(
        string rootPath,
        List<string> findValues,
        bool isMulti = true,
        List<string> returnColumnNames = null,
        string searchColumnName = "B"
    )
    {
        var sheetNames = MiniExcel.GetSheetNames(rootPath);

        var targetList = new Dictionary<string, List<string>>();

        foreach (var findValue in findValues)
        {
            foreach (var sheetName in sheetNames)
            {
                if (sheetName.Contains("#") || sheetName.Contains("Sheet") && sheetName != "Sheet1")
                    continue;
                var rows = MiniExcel
                    .Query(rootPath, sheetName: sheetName)
                    .Cast<IDictionary<string, object>>();
                var result = rows.FirstOrDefault(row =>
                    row.ContainsKey(searchColumnName)
                    && row[searchColumnName]?.ToString() == findValue
                );

                if (result == null)
                    continue;

                var returnValue = new List<string>();
                foreach (var returnColumnName in returnColumnNames)
                {
                    returnValue.Add(result[returnColumnName].ToString());
                }

                targetList[findValue] = returnValue;
            }
        }

        return targetList;
    }

    //MiniExcel查询：全局或指定查询
    public static Dictionary<string, List<string>> SearchModelKeyMiniExcel(
        string findValue,
        string[] files,
        bool isFixList,
        bool isMulti
    )
    {
        var targetList = isMulti
            ? (IDictionary<string, List<string>>)new ConcurrentDictionary<string, List<string>>()
            : new Dictionary<string, List<string>>();
        var currentCount = 0;
        var count = files.Length;
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");

        Action<string> processFile = file =>
        {
            var fileName = Path.GetFileName(file);
            var sheetFullName = fileName;

            try
            {
                var sheetNames = MiniExcel.GetSheetNames(file);

                foreach (var sheetName in sheetNames)
                {
                    if (sheetName.Contains("#") || sheetName.Contains("Chart"))
                        continue;

                    var rows = MiniExcel.Query(
                        file,
                        sheetName: sheetName,
                        startCell: "A2",
                        useHeaderRow: true
                    );

                    sheetFullName = fileName.Contains("$") ? $"{fileName}#{sheetName}" : fileName;

                    int rowIndex = 1;
                    foreach (var row in rows)
                    {
                        int colIndex = 1;
                        foreach (var cell in row)
                        {
                            // 只搜索第2列
                            if (colIndex == 2)
                            {
                                var cellValue = cell.Value?.ToString();
                                if (
                                    cellValue != null
                                    && (
                                        isAll
                                            ? cellValue.StartsWith(findValue)
                                            : cellValue == findValue
                                    )
                                )
                                {
                                    // 确保 targetList 中存在 fileName 键
                                    if (!targetList.ContainsKey(sheetFullName))
                                    {
                                        targetList[sheetFullName] = new List<string>();
                                    }

                                    targetList[sheetFullName].Add(cellValue);
                                }
                            }

                            colIndex++;
                        }

                        rowIndex++;
                    }

                    if (targetList.ContainsKey(sheetFullName) && isFixList)
                    {
                        var list = targetList[sheetFullName];
                        if (list.Count > 1)
                        {
                            targetList[sheetFullName] = new List<string>
                            {
                                list.First(),
                                list.Last()
                            };
                        }
                        else if (list.Count == 1)
                        {
                            list.Add(list.First());
                        }
                    }
                }
            }
            catch
            {
                // 记录异常信息，继续处理下一个文件
            }

            Interlocked.Increment(ref currentCount);
            NumDesAddIn.App.StatusBar = $"正在检查第 {currentCount}/{count} 个文件: {file}";
        };

        if (isMulti)
        {
            var options = new ParallelOptions
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };
            Parallel.ForEach(files, options, processFile);
        }
        else
        {
            foreach (var file in files)
            {
                processFile(file);
            }
        }

        return new Dictionary<string, List<string>>(targetList);
    }

    //MiniExcel查询：全局查询Sheet名
    public static List<(string, string, int, int)> SearchSheetNameFromExcel(
        string rootPath,
        string findValue,
        bool isMulti = false
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new ConcurrentBag<(string, string, int, int)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");

        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Action<string> processFile = file =>
        {
            try
            {
                var sheetNames = MiniExcel.GetSheetNames(file);
                var fileName = Path.GetFileName(file);

                bool isFileMatch = isAll
                    ? fileName.Contains(findValue + ".xlsx")
                    : fileName == findValue + ".xlsx";

                bool isSheetMatch(string sheetName) =>
                    isAll ? sheetName.Contains(findValue) : sheetName == findValue;

                if (isFileMatch)
                {
                    targetList.Add((file, sheetNames[0], 1, 1));
                }
                else
                {
                    foreach (var sheetName in sheetNames)
                    {
                        if (sheetName.Contains("#"))
                            continue;

                        if (isSheetMatch(sheetName))
                        {
                            targetList.Add((file, sheetName, 1, 1));
                        }
                    }
                }
            }
            catch
            {
                // 记录异常信息，继续处理下一个文件
            }
        };

        if (isMulti)
        {
            Parallel.ForEach(files, options, processFile);
        }
        else
        {
            foreach (var file in files)
            {
                processFile(file);
            }
        }

        return targetList.ToList();
    }

    public static List<(string, string, int, int)> SearchFormularNameFromExcel(string findValue)
    {
        var wkPath = Wk.FullName;
        var targetList = new List<(string, string, int, int)>();

        using (var package = new ExcelPackage(new FileInfo(wkPath)))
        {
            var sheetCount = package.Workbook.Worksheets.Count;
            for (int i = 0; i < sheetCount; i++)
            {
                var sheet = package.Workbook.Worksheets[i];
                var formulaCells = sheet
                    .Cells.Where(c => c.Formula != null && c.Formula.Contains(findValue))
                    .Select(c => new
                    {
                        Address = c.Address,
                        Formula = c.Formula,
                        Row = c.Start.Row, // 获取行号（基于1）
                        Column = c.Start.Column // 获取列号（基于1）
                    });

                foreach (var cell in formulaCells)
                {
                    targetList.Add((wkPath, sheet.Name, cell.Row, cell.Column));
                }
            }
        }

        return targetList.ToList();
    }
    #endregion
}
