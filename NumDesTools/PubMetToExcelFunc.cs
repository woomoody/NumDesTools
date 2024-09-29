using System.Collections.Concurrent;
using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using MiniExcelLibs;
using NLua;
using NPOI.SS.UserModel;
using NumDesTools.UI;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using MessageBox = System.Windows.MessageBox;
using Process = System.Diagnostics.Process;


// ReSharper disable All

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类调用的具体业务逻辑
/// </summary>
public class PubMetToExcelFunc
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
        var sheet = NumDesAddIn.App.ActiveSheet;
        var selectCell = NumDesAddIn.App.ActiveCell;
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var workBookName = workBook.Name;
        var workbookPath = workBook.Path;
        var sheetName = sheet.Name;

        if (workbookPath!.Contains("克朗代克"))
        {
            workBookName = "克朗代克##" + workBookName;
        }

        workbookPath = Path.GetDirectoryName(workbookPath);

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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

    #region Excel数据查找

    //Epplus
    public static List<(string, string, int, int)> SearchKeyFromExcel(
        string rootPath,
        string errorValue
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var filesCollection = new SelfExcelFileCollector(rootPath, 2);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new List<(string, string, int, int)>();
        var currentCount = 0;
        var count = files.Length;
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");
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
                                            ? cellValue.Contains(errorValue)
                                            : cellValue == errorValue
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
        string errorValue
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var filesCollection = new SelfExcelFileCollector(rootPath, 2);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new List<(string, string, int, int)>();
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");
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
                                                ? cellValue.Contains(errorValue)
                                                : cellValue == errorValue
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

    //MiniExcel
    public static List<(string, string, int, int)> SearchKeyFromExcelMiniExcel(
        string rootPath,
        string errorValue
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath, 2);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new List<(string, string, int, int)>();
        var currentCount = 0;
        var count = files.Length;
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");

        foreach (var file in files)
        {
            try
            {
                var sheetNames = MiniExcel.GetSheetNames(file);
                foreach (var sheetName in sheetNames)
                {
                    if (sheetName.Contains("#"))
                        continue;

                    var rows = MiniExcel.Query(file, sheetName: sheetName);
                    int rowIndex = 1;
                    foreach (var row in rows)
                    {
                        int colIndex = 1;
                        foreach (var cell in row)
                        {
                            var cellValue = cell.Value?.ToString();
                            if (
                                cellValue != null
                                && (
                                    isAll ? cellValue.Contains(errorValue) : cellValue == errorValue
                                )
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

            currentCount++;
            NumDesAddIn.App.StatusBar = $"正在检查第 {currentCount}/{count} 个文件: {file}";
        }

        return targetList;
    }

    public static List<(string, string, int, int)> SearchKeyFromExcelMultiMiniExcel(
        string rootPath,
        string errorValue
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath, 2);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new ConcurrentBag<(string, string, int, int)>();
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");

        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Parallel.ForEach(
            files,
            options,
            file =>
            {
                try
                {
                    var sheetNames = MiniExcel.GetSheetNames(file);
                    Parallel.ForEach(
                        sheetNames,
                        sheetName =>
                        {
                            if (sheetName.Contains("#"))
                                return;

                            var rows = MiniExcel.Query(file, sheetName: sheetName);
                            int rowIndex = 1;
                            foreach (var row in rows)
                            {
                                int colIndex = 1;
                                foreach (var cell in row)
                                {
                                    var cellValue = cell.Value?.ToString();
                                    if (
                                        cellValue != null
                                        && (
                                            isAll
                                                ? cellValue.Contains(errorValue)
                                                : cellValue == errorValue
                                        )
                                    )
                                    {
                                        targetList.Add((file, sheetName, rowIndex, colIndex));
                                    }

                                    colIndex++;
                                }

                                rowIndex++;
                            }
                        }
                    );
                }
                catch
                {
                    // 记录异常信息，继续处理下一个文件
                }
            }
        );

        return targetList.ToList();
    }

    public static List<(string, string, int, int)> SearchKeyFromExcelIDMultiMiniExcel(
        string rootPath,
        string errorValue
    )
    {
        var filesCollection = new SelfExcelFileCollector(rootPath, 2);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new ConcurrentBag<(string, string, int, int)>();
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");

        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Parallel.ForEach(
            files,
            options,
            file =>
            {
                try
                {
                    var sheetNames = MiniExcel.GetSheetNames(file);
                    Parallel.ForEach(
                        sheetNames,
                        sheetName =>
                        {
                            if (sheetName.Contains("#"))
                                return;

                            var rows = MiniExcel.Query(file, sheetName: sheetName);
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
                                                    ? cellValue.Contains(errorValue)
                                                    : cellValue == errorValue
                                            )
                                        )
                                        {
                                            targetList.Add((file, sheetName, rowIndex, colIndex));
                                        }

                                        break;
                                    }

                                    colIndex++;
                                }

                                rowIndex++;
                            }
                        }
                    );
                }
                catch
                {
                    // 记录异常信息，继续处理下一个文件
                }
            }
        );

        return targetList.ToList();
    }
    public static Dictionary<string, List<string>> SearchModelKeyFromExcelMiniExcel(
string rootPath,
string errorValue
)
    {
        var filesCollection = new SelfExcelFileCollector(rootPath, 2);
        var files = filesCollection.GetAllExcelFilesPath();

        var targetList = new Dictionary<string, List<string>>();

        var currentCount = 0;
        var count = files.Length;
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");

        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

        Parallel.ForEach(
            files,
            options,
            file =>
            {
                var fileName = Path.GetFileName(file);
                try
                {
                    var sheetNames = MiniExcel.GetSheetNames(file);
                    Parallel.ForEach(
                        sheetNames,
                        sheetName =>
                        {
                            if (sheetName.Contains("#"))
                                return;

                            var rows = MiniExcel.Query(file, sheetName: sheetName);
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
                                                    ? cellValue.StartsWith(errorValue)
                                                    : cellValue == errorValue
                                            )
                                        )
                                        {
                                            if (!sheetName.Contains("#"))
                                            {
                                                if (sheetName.Contains("Sheet") || sheetName.Contains("map_proto"))
                                                {
                                                    fileName = $"{fileName}";
                                                }
                                                else
                                                {
                                                    fileName = $"{fileName}#{sheetName}";
                                                }

                                                // 确保 targetList 中存在 fileName 键
                                                if (!targetList.ContainsKey(fileName))
                                                {
                                                    targetList[fileName] = new List<string>();
                                                }

                                                targetList[fileName].Add(cellValue);
                                            }

                                        }
                                    }

                                    colIndex++;
                                }

                                rowIndex++;
                            }
                        }
                    );
                }
                catch
                {
                    // 记录异常信息，继续处理下一个文件
                }

                //处理List
                if (targetList.ContainsKey(fileName))
                {
                    if (targetList[fileName].Count > 1)
                    {
                        // 只保留第一个和最后一个元素
                        var first = targetList[fileName].First();
                        var last = targetList[fileName].Last();
                        targetList[fileName] = new List<string> { first, last };
                    }
                    else if (targetList[fileName].Count == 1)
                    {
                        // 复制第一个元素并添加到第二个位置
                        var first = targetList[fileName].First();
                        targetList[fileName].Add(first);
                    }
                }

                currentCount++;
                NumDesAddIn.App.StatusBar = $"正在检查第 {currentCount}/{count} 个文件: {file}";
            }
        );

        return targetList;
    }
    #endregion

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

    //遍历目录文件
    public static void ExcelFolderPath(string[] folder)
    {
        var baseFolder = folder[1];
        var newPath = Path.GetDirectoryName(
            Path.GetDirectoryName(Path.GetDirectoryName(baseFolder))
        );
        if (newPath != null)
        {
            var filesCollection = new SelfExcelFileCollector(newPath, 2);
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
        }

        var needFixLinks = new List<string>();
        foreach (string link in links)
        {
            if (!baseFilePathList.Contains(link))
            {
                var fileName = Path.GetFileName(link);
                var filePath = Path.GetDirectoryName(link);
                var newLink = filePath + @"\[" + fileName + @"]";
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
                            string originalFormula = cell.Formula;
                            string newFormula = originalFormula;
                            for (int indexFor = 0; indexFor < needFixLinks.Count; indexFor++)
                            {
                                var oldFor = needFixLinks[indexFor];
                                var newFor = replaceFixLinks[indexFor];
                                // 替换公式样式
                                newFormula = newFormula.Replace(oldFor, newFor);
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
        var filesCollector = new SelfExcelFileCollector(rootPath, 2);
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
        var newFolderPath = WkPath+@"\";
        dynamic queries = Wk.GetType().InvokeMember("Queries", System.Reflection.BindingFlags.GetProperty, null, Wk, null);
        foreach (dynamic query in queries)
        {
            string name = query.Name;
            string formula = query.Formula;
            string oldFilePath = ExtractFilePathFromFormula(formula);
            if (!string.IsNullOrEmpty(oldFilePath))
            {
                string fileName = System.IO.Path.GetFileName(oldFilePath);
                string newFilePath = System.IO.Path.Combine(newFolderPath, fileName);
                string newFormula = formula.Replace(oldFilePath, newFilePath);
                query.GetType().InvokeMember("Formula", System.Reflection.BindingFlags.SetProperty, null, query, new object[] { newFormula });
            }
        }
    }
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
    public class ExcelDataFormatCheck
    {
        public static void CheckValueFormat(string specialCharsStr)
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
                    for (int i = 0 ; i < specialChars.Length; i++) 
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
    }
}
