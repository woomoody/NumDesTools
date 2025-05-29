using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Match = System.Text.RegularExpressions.Match;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416


namespace NumDesTools;

/// <summary>
/// Merge项目Excel文件数据自动处理类集合
/// </summary>
public static class ExcelDataAutoInsert
{
    [ExcelFunction(IsHidden = true)]
    public static int FindTitle(dynamic sheet, int rows, string findValue)
    {
        var maxColumn = sheet.UsedRange.Columns.Count;
        for (var column = 1; column <= maxColumn; column++)
            if (sheet.Cells[rows, column] is Range cell && cell.Value2?.ToString() == findValue)
                return column;
        return -1;
    }

    [ExcelFunction(IsHidden = true)]
    public static string ErrorExcelMark(dynamic errorExcelList, dynamic sheet)
    {
        var strBuild = new StringBuilder();
        for (var i = 0; i < errorExcelList.Count; i++)
        {
            if (errorExcelList[i][0].Item1 == 0)
                continue;
            strBuild.Append(errorExcelList[i][0].Item2);
            var cell = sheet.Cells[errorExcelList[i][0].Item1, 1];
            cell.Value = "git checkout -- Excels/Tables/" + errorExcelList[i][0].Item3;
            cell.Font.Color = Color.Red;
        }

        var errorLog = strBuild.ToString();
        return errorLog;
    }

    public static string StringRegPlace(string str, List<(int, int)> digit, int addValue)
    {
        var reg = "\\d+";
        var matches = Regex.Matches(str, reg);
        var matchCount = 0;
        var digitCount = 0;
        foreach (Match unused in matches)
        {
            var matches2 = Regex.Matches(str, reg);
            var match2 = matches2[matchCount];
            var numStr = match2.Value;
            var index = match2.Index;
#pragma warning disable CA1305
            var num = long.Parse(numStr);
#pragma warning restore CA1305
            if (digit.Any(item => item.Item1 == matchCount + 1))
            {
                var addDigit = (long)Math.Pow(10, digit[digitCount].Item2 - 1) * addValue;
                var newNum = num + addDigit;
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
                digitCount++;
            }
            else if (digit is [{ Item1: 0 } _])
            {
                if (digit[0].Item2 > 1000)
                {
                    str = "^error^";
                    return str;
                }

                var addDigit = Math.Abs((long)Math.Pow(10, digit[0].Item2 - 1) * addValue);
                if (addDigit > (num + 1) * 100)
                {
                    str = "^error^";
                    return str;
                }

                var newNum = num + addDigit;
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
            }

            matchCount++;
        }

        return str;
    }

    public static void ExcelHyperLinks(dynamic excelPath, dynamic sheet)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        for (var i = 2; i <= 500; i++)
        {
            var modeCol = FindTitle(sheet, 1, "实际模板(上一期)");
            var excelName = FindTitle(sheet, 1, "表名");
            string findValue = sheet.Cells[i, modeCol].Value?.ToString();
            var cell = sheet.Cells[i, excelName];
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx"))
                continue;
            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
            string path = cell.value switch
            {
                "Localizations.xlsx" => newPath + @"\Excels\Localizations\Localizations.xlsx",
                "UIConfigs.xlsx" => newPath + @"\Excels\UIs\UIConfigs.xlsx",
                "UIItemConfigs.xlsx" => newPath + @"\Excels\UIs\UIItemConfigs.xlsx",
                _ => excelPath + @"\" + cell.value
            };

            var excel = new ExcelPackage(new FileInfo(path));
            var workbook = excel.Workbook;
            var sheetTemp = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];
            var row = PubMetToExcel.FindSourceRow(sheetTemp, 2, findValue);
            if (row != 0)
            {
                var newRow = "A" + row;

                var sheetName = sheetTemp.Name;
                var links = path + "#" + sheetName + "!" + newRow;
                excel.Dispose();
                cell.Hyperlinks.Add(cell, links);
                cell.Font.Size = 9;
                cell.Font.Name = "微软雅黑";
            }
        }
    }

    public static void ExcelHyperLinksNormal(dynamic excelPath, dynamic sheet)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        for (var i = 2; i <= 500; i++)
        {
            var cell = sheet.Cells[i, 5];
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx"))
                continue;
            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
            string path = cell.value switch
            {
                "Localizations.xlsx" => newPath + @"\Excels\Localizations\Localizations.xlsx",
                "UIConfigs.xlsx" => newPath + @"\Excels\UIs\UIConfigs.xlsx",
                "UIItemConfigs.xlsx" => newPath + @"\Excels\UIs\UIItemConfigs.xlsx",
                _ => excelPath + @"\" + cell.value
            };

            var links = path + "#" + "Sheet1!A1";
            cell.Hyperlinks.Add(cell, links);
            cell.Font.Size = 9;
            cell.Font.Name = "微软雅黑";
        }
    }

    public static List<(int, int)> CellFixValueKeyList(string str)
    {
        var monkeyList = new List<(int, int)>();

        str ??= "";

        if (str.Contains(','))
        {
            var pairs = str.Split(',');
            foreach (var pair in pairs)
                if (pair.Contains('#'))
                {
                    var parts = pair.Split('#');
                    if (!int.TryParse(parts[0], out var key))
                    {
                        MessageBox.Show($@"{str}#前必须有数值");
                        Environment.Exit(0);
                    }

                    if (!int.TryParse(parts[1], out var value))
                        value = 1;

                    monkeyList.Add((key, value));
                }
                else
                {
#pragma warning disable CA1305
                    monkeyList.Add((int.Parse(pair), 1));
#pragma warning restore CA1305
                }
        }
        else
        {
            if (str.Contains('#'))
            {
                var parts = str.Split('#');
#pragma warning disable CA1305
                var key = Convert.ToInt32(parts[0]);
#pragma warning restore CA1305
#pragma warning disable CA1305
                var value = Convert.ToInt32(parts[1]);
#pragma warning restore CA1305
                monkeyList.Add((key, value));
            }
            else
            {
                int strTemp;
                if (str == "")
                {
                    strTemp = 0;
                    monkeyList.Add((strTemp, 1));
                }
                else
                {
#pragma warning disable CA1305
                    strTemp = int.Parse(str);
#pragma warning restore CA1305
                    monkeyList.Add((0, strTemp));
                }
            }
        }

        return monkeyList;
    }

    [ExcelFunction(IsHidden = true)]
    public static string ExcelPathIgnore(dynamic excelPath, dynamic excelName)
    {
        string path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        switch (excelName)
        {
            case "Localizations.xlsx":
                path = newPath + @"\Excels\Localizations\Localizations.xlsx";
                break;
            case "UIConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIConfigs.xlsx";
                break;
            case "UIItemConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIItemConfigs.xlsx";
                break;
            default:
                path = excelPath + @"\" + excelName;
                break;
        }

        return path;
    }
}

public static class ExcelDataAutoInsertLanguage
{
    public static void AutoInsertData()
    {
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var excelPath = workBook.Path;
        var sourceSheet = workBook.Worksheets["多语言对话【模板】"];
        var fixSheet = workBook.Worksheets["数据修改"];
        var classSheet = workBook.Worksheets["枚举数据"];
        var emoSheet = workBook.Worksheets["表情枚举"];

        ErrorLogCtp.DisposeCtp();

        var errorExcelList = new List<List<(int, string, string)>>();
        if (errorExcelList == null)
            throw new ArgumentNullException(nameof(errorExcelList));

        List<(int, string, string)> error = LanguageDialogData(
            sourceSheet,
            fixSheet,
            classSheet,
            emoSheet,
            excelPath,
            NumDesAddIn.App
        );

        if (error.Count != 0)
            errorExcelList.Add(error);

        string errorLog = ExcelDataAutoInsert.ErrorExcelMark(errorExcelList, fixSheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }

        Marshal.ReleaseComObject(sourceSheet);
        Marshal.ReleaseComObject(fixSheet);
        Marshal.ReleaseComObject(classSheet);
        Marshal.ReleaseComObject(emoSheet);
        Marshal.ReleaseComObject(workBook);
    }

    private static List<(int, string, string)> LanguageDialogData(
        dynamic sourceSheet,
        dynamic fixSheet,
        dynamic classSheet,
        dynamic emoSheet,
        string excelPath,
        dynamic app
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var sourceData = PubMetToExcel.ExcelDataToList(sourceSheet);

        var sourceTitle = sourceData.Item1;
        var sourceDataList = sourceData.Item2;

        var fixData = PubMetToExcel.ExcelDataToList(fixSheet);
        var fixTitle = fixData.Item1;
        var fixDataList = fixData.Item2;

        var classData = PubMetToExcel.ExcelDataToList(classSheet);
        var classTitle = classData.Item1;
        var classDataList = classData.Item2;

        var emoData = PubMetToExcel.ExcelDataToList(emoSheet);
        var emoDataList = emoData.Item2;

        var fileIndex = fixTitle.IndexOf("表名");
        var keyIndex = fixTitle.IndexOf("字段");
        var modelIdIndex = fixTitle.IndexOf("初始模板");

        var errorExcel = 0;
        var errorList = new List<(int, string, string)>();

        for (var i = 0; i < fixDataList.Count; i++)
        {
            if (fixDataList[i][fileIndex] == null)
                continue;
            var sourceKeyList = new List<string>();
            var fixKeyList = new List<string>();
            for (int j = keyIndex; j < fixTitle.Count; j++)
            {
                if (fixDataList[i][j] != null)
                {
                    var sourceKey = fixDataList[i][j].ToString();
                    sourceKeyList.Add(sourceKey);
                }

                if (fixDataList[i + 1][j] != null)
                {
                    var fixKey = fixDataList[i + 1][j].ToString();
                    fixKeyList.Add(fixKey);
                }
            }

            var fixFileName = fixDataList[i][fileIndex].ToString();
            var fixFileModeId = fixDataList[i][modelIdIndex].ToString();

            string path = ExcelDataAutoInsert.ExcelPathIgnore(excelPath, fixFileName);
            var targetExcel = new ExcelPackage(new FileInfo(path));
            ExcelWorkbook targetBook;
            string errorExcelLog;
            try
            {
                targetBook = targetExcel.Workbook;
            }
            catch (Exception ex)
            {
                errorExcel = i * 2 + 2;
                errorExcelLog = fixFileName + "#不能创建WorkBook对象" + ex.Message;
                errorList.Add((errorExcel, errorExcelLog, fixFileName));
                continue;
            }

            ExcelWorksheet targetSheet;
            try
            {
                targetSheet = targetBook.Worksheets["Sheet1"] ?? targetBook.Worksheets[0];
            }
            catch (Exception ex)
            {
                errorExcel = i * 2 + 2;
                errorExcelLog = fixFileName + "#不能创建WorkBook对象" + ex.Message;
                errorList.Add((errorExcel, errorExcelLog, fixFileName));
                continue;
            }

            var c = 0;
            if (fixFileName == "GuideDialogDetail.xlsx")
                c = 1;
            else if (fixFileName == "Localizations.xlsx")
                c = 2;
            else if (fixFileName == "GuideDialogBranch.xlsx")
                c = 3;
            else if (fixFileName == "GuideDialogDetailSpecialSetting.xlsx")
                c = 4;
            var idList = new List<string>();
            for (var r = 0; r < sourceDataList.Count; r++)
            {
                var value = sourceDataList[r][c]?.ToString() ?? "";
                idList.Add(value);
            }

            var newIdList = idList.Distinct().ToList();

            var rowsToDelete = new List<int>();
            foreach (var id in newIdList)
            {
                var reDd = PubMetToExcel.FindSourceRow(targetSheet, 2, id);
                if (reDd != -1)
                    rowsToDelete.Add(reDd);
            }

            rowsToDelete.Sort();
            rowsToDelete.Reverse();
            foreach (var rowToDelete in rowsToDelete)
                try
                {
                    targetSheet.DeleteRow(rowToDelete, 1);
                }
                catch (Exception ex)
                {
                    errorExcel = i * 2 + 2;
                    errorExcelLog = fixFileName + "无法删除重复数据，sheet格式问题，复制数据到新表" + ex.Message;
                    errorList.Add((errorExcel, errorExcelLog, fixFileName));
                }

            var endRowSource = PubMetToExcel.FindSourceRow(targetSheet, 2, fixFileModeId);
            if (endRowSource == -1)
            {
                MessageBox.Show(fixFileModeId + @"目标表中不存在");
                continue;
            }

            targetSheet.InsertRow(endRowSource + 1, sourceDataList.Count);
            var colCount = targetSheet.Dimension.Columns;
            //多语言表不需要复制全部列
            if (fixFileName == "Localizations.xlsx")
            {
                colCount = 7;
            }
            var cellSource = targetSheet.Cells[endRowSource, 1, endRowSource, colCount];
            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var cellTarget = targetSheet.Cells[
                    endRowSource + 1 + m,
                    1,
                    endRowSource + 1 + m,
                    colCount
                ];
                cellSource.Copy(
                    cellTarget,
                    ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting
                        | ExcelRangeCopyOptionFlags.ExcludeMergedCells
                );
                cellSource.CopyStyles(cellTarget);
            }

            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var errorCell = sourceDataList[m][0];
                if (errorCell == null)
                {
                    errorExcelLog = sourceSheet.Name + "#表格尾行有，多余格式，清除";
                    MessageBox.Show(errorExcelLog);
                    return errorList;
                }

                var sourceCount = 0;
                foreach (var source in sourceKeyList)
                {
                    var cellCol = PubMetToExcel.FindSourceCol(
                        targetSheet,
                        2,
                        fixKeyList[sourceCount]
                    );
                    if (cellCol == -1)
                    {
                        if (fixKeyList[sourceCount] == "bgType")
                        {
                            sourceCount++;
                            continue;
                        }

                        errorExcel = i * 2 + 2;
                        errorExcelLog = fixFileName + "#表格字段#[" + fixKeyList[sourceCount] + "]未找到";
                        errorList.Add((errorExcel, errorExcelLog, fixFileName));
                        sourceCount++;
                        continue;
                    }

                    var cellTarget = targetSheet.Cells[endRowSource + 1 + m, cellCol];
                    var newStr = "";
                    if (int.TryParse(source, out var e))
                    {
                        var realCol = "";
                        if (fixFileName == "GuideDialogGroup.xlsx")
                            realCol = "GroupID";
                        else if (fixFileName == "GuideDialogBranch.xlsx")
                            realCol = "BranchID";
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(realCol)];
#pragma warning disable CS0252
                        if (sourceValue == "" || sourceValue == null)
                            continue;
#pragma warning restore CS0252
                        var str = sourceValue.ToString();
                        var digit = Math.Pow(10, e);
                        var repeatCount = 0;
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf(realCol)];
#pragma warning disable CS0252
                            if (repeatValue == "" || repeatValue == null)
                                continue;
#pragma warning restore CS0252
                            if (repeatValue == sourceValue)
                            {
#pragma warning disable CA1305
                                var newNum = long.Parse(str) * digit + repeatCount + 1;
#pragma warning restore CA1305
                                newStr = newStr + newNum + ",";
                                repeatCount++;
                            }
                        }

                        newStr = "[" + newStr.Substring(0, newStr.Length - 1) + "]";
                        cellTarget.Value = newStr;
                    }
                    else if (source == "枚举1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var scCol = classTitle.IndexOf(source);
                        var newId = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newId = classDataList[k][scCol + 1].ToString();
                                break;
                            }
                        }

                        var sourceStr = cellTarget.Value?.ToString();
                        var reg = "\\d+";
                        if (sourceStr == null || sourceStr == "")
                            continue;
                        var matches = Regex.Matches(sourceStr, reg);

                        var oldId = matches[0].Value.ToString();
                        if (newId != "")
                            sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色表情")
                    {
                        var sourceValue = sourceDataList[m]
                            [sourceTitle.IndexOf(source)]
                            ?.ToString();
                        for (var k = 0; k < emoDataList.Count; k++)
                        {
                            var targetValue = emoDataList[k][0].ToString();
                            if (targetValue == sourceValue)
                            {
                                var emoId = emoDataList[k][2];
                                if (emoId == null)
                                {
                                    emoId = "idle";
                                }
                                cellTarget.Value = emoId;
                                break;
                            }
                        }
                    }
                    else if (source == "触发分支")
                    {
                        var sourceValue = sourceDataList[m]
                            [sourceTitle.IndexOf(source)]
                            ?.ToString();
                        if (sourceValue == null || sourceValue == "" || sourceValue == "0")
                        {
                            sourceCount++;
                            continue;
                        }

                        var uniqueValues1 = new HashSet<string>();
                        var strBranch = "";
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k]
                                [sourceTitle.IndexOf("分支归属")]
                                ?.ToString();
                            if (repeatValue == null || repeatValue == "")
                                continue;
                            if (repeatValue == sourceValue)
                            {
                                var branchId = sourceDataList[k][sourceTitle.IndexOf("BranchID")];
                                if (uniqueValues1.Add((string)branchId))
                                    strBranch = strBranch + branchId + ",";
                            }
                        }

                        strBranch = "[" + strBranch.Substring(0, strBranch.Length - 1) + "]";
                        cellTarget.Value = strBranch;
                    }
                    //else if (source == "分支多语言")
                    //{
                    //    var newId = sourceDataList[m][sourceTitle.IndexOf("BranchID")]?.ToString();
                    //    var sourceStr = cellTarget.Value?.ToString();
                    //    if (sourceStr == null || sourceStr == "")
                    //        continue;
                    //    var reg = "\\d+";
                    //    var matches = Regex.Matches(sourceStr, reg);
                    //    var oldId = matches[0].Value.ToString();
                    //    if (newId != "")
                    //        sourceStr = sourceStr.Replace(oldId, newId);
                    //    cellTarget.Value = sourceStr;
                    //}
                    else if (source == "角色换装1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m]
                            [sourceTitle.IndexOf("角色换装")]
                            ?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue =
                                    sourceValue2 == "1"
                                        ? (string)classDataList[k][scCol + 2].ToString()
                                        : "[]";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "角色换装2")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m]
                            [sourceTitle.IndexOf("角色换装")]
                            ?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue =
                                    sourceValue2 != "1"
                                        ? (string)classDataList[k][scCol + 3].ToString()
                                        : "";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "UI对话框")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("UI对话框")];
                        sourceValue = sourceValue == null ? "1" : sourceValue.ToString();
                        if (fixKeyList[sourceCount] == "bgType")
                            cellTarget.Value = sourceValue;
                    }
                    else
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)];
                        cellTarget.Value = sourceValue;
                    }

                    sourceCount++;
                }
            }

            if (errorExcel != 0)
                continue;
            int startRow = endRowSource + 1;
            int endRow2 = startRow + sourceDataList.Count - 1;
            if (
                fixFileName == "GuideDialogBranch.xlsx"
                || fixFileName == "GuideDialogGroup.xlsx"
                || fixFileName == "GuideDialogDetailSpecialSetting.xlsx"
                || fixFileName == "Localizations.xlsx"
            )
            {
                var uniqueValues = new HashSet<string>();
                for (var row = 4; row <= endRow2; row++)
                {
                    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";

                    if (uniqueValues.Contains(cellValue) || string.IsNullOrWhiteSpace(cellValue))
                    {
                        targetSheet.DeleteRow(row);
                        row--;
                        endRow2--;
                    }
                    else
                    {
                        uniqueValues.Add(cellValue);
                    }
                }
            }

            targetExcel.Save();
            targetExcel.Dispose();
            var excelCount = i / 2 + 1;
            app.StatusBar =
                "写入数据" + "<" + excelCount + "/" + fixDataList.Count / 2 + ">" + fixFileName;
        }

        return errorList;
    }

    public static void AutoInsertDataByUd(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var excelPath = workBook.Path;
        var sourceSheet = workBook.Worksheets["多语言对话【模板】"];
        var fixSheet = workBook.Worksheets["数据修改"];
        var classSheet = workBook.Worksheets["枚举数据"];
        var emoSheet = workBook.Worksheets["表情枚举"];

        ErrorLogCtp.DisposeCtp();

        var errorExcelList = new List<List<(int, string, string)>>();
        if (errorExcelList == null)
            throw new ArgumentNullException(nameof(errorExcelList));

        List<(int, string, string)> error = LanguageDialogDataByUd(
            sourceSheet,
            fixSheet,
            classSheet,
            emoSheet,
            excelPath,
            NumDesAddIn.App
        );

        if (error.Count != 0)
            errorExcelList.Add(error);

        string errorLog = ExcelDataAutoInsert.ErrorExcelMark(errorExcelList, fixSheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }

        NumDesAddIn.App.StatusBar = "导出完成";
        Marshal.ReleaseComObject(sourceSheet);
        Marshal.ReleaseComObject(fixSheet);
        Marshal.ReleaseComObject(classSheet);
        Marshal.ReleaseComObject(emoSheet);
        Marshal.ReleaseComObject(workBook);
    }

    private static List<(int, string, string)> LanguageDialogDataByUd(
        dynamic sourceSheet,
        dynamic fixSheet,
        dynamic classSheet,
        dynamic emoSheet,
        string excelPath,
        dynamic app
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var sourceData = PubMetToExcel.ExcelDataToListBySelfToEnd(sourceSheet, 0, 1, 1);
        var sourceTitle = sourceData.Item1;
        var sourceDataList = sourceData.Item2;

        var fixData = PubMetToExcel.ExcelDataToList(fixSheet);
        var fixTitle = fixData.Item1;
        var fixDataList = fixData.Item2;

        var classData = PubMetToExcel.ExcelDataToList(classSheet);
        var classTitle = classData.Item1;
        var classDataList = classData.Item2;

        var emoData = PubMetToExcel.ExcelDataToList(emoSheet);
        var emoDataList = emoData.Item2;

        var fileIndex = fixTitle.IndexOf("表名");
        var keyIndex = fixTitle.IndexOf("字段");

        var errorExcel = 0;
        var errorList = new List<(int, string, string)>();

        for (var i = 0; i < fixDataList.Count; i++)
        {
            if (fixDataList[i][fileIndex] == null)
                continue;
            var sourceKeyList = new List<string>();
            var fixKeyList = new List<string>();
            for (int j = keyIndex; j < fixTitle.Count; j++)
            {
                if (fixDataList[i][j] != null)
                {
                    var sourceKey = fixDataList[i][j].ToString();
                    sourceKeyList.Add(sourceKey);
                }

                if (fixDataList[i + 1][j] != null)
                {
                    var fixKey = fixDataList[i + 1][j].ToString();
                    fixKeyList.Add(fixKey);
                }
            }

            var fixFileName = fixDataList[i][fileIndex].ToString();

            string path = ExcelDataAutoInsert.ExcelPathIgnore(excelPath, fixFileName);
            var targetExcel = new ExcelPackage(new FileInfo(path));
            ExcelWorkbook targetBook;
            string errorExcelLog;
            try
            {
                targetBook = targetExcel.Workbook;
            }
            catch (Exception ex)
            {
                errorExcel = i * 2 + 2;
                errorExcelLog = fixFileName + "#不能创建WorkBook对象" + ex.Message;
                errorList.Add((errorExcel, errorExcelLog, fixFileName));
                continue;
            }

            ExcelWorksheet targetSheet;
            try
            {
                targetSheet = targetBook.Worksheets["Sheet1"] ?? targetBook.Worksheets[0];
            }
            catch (Exception ex)
            {
                errorExcel = i * 2 + 2;
                errorExcelLog = fixFileName + "#不能创建WorkBook对象" + ex.Message;
                errorList.Add((errorExcel, errorExcelLog, fixFileName));
                continue;
            }

            var c = 0;
            if (fixFileName == "GuideDialogDetail.xlsx")
                c = 1;
            else if (fixFileName == "Localizations.xlsx")
                c = 2;
            else if (fixFileName == "GuideDialogBranch.xlsx")
                c = 3;
            else if (fixFileName == "GuideDialogDetailSpecialSetting.xlsx")
                c = 4;
            var idList = new List<string>();
            for (var r = 0; r < sourceDataList.Count; r++)
            {
                var value = sourceDataList[r][c]?.ToString() ?? "";
                idList.Add(value);
            }

            var newIdList = idList.Distinct().ToList();

            var rowsToDelete = new List<int>();
            foreach (var id in newIdList)
            {
                var reDd = PubMetToExcel.FindSourceRow(targetSheet, 2, id);
                if (reDd != -1)
                    rowsToDelete.Add(reDd);
            }

            rowsToDelete.Sort();
            rowsToDelete.Reverse();
            foreach (var rowToDelete in rowsToDelete)
                try
                {
                    targetSheet.DeleteRow(rowToDelete, 1);
                }
                catch (Exception ex)
                {
                    errorExcel = i * 2 + 2;
                    errorExcelLog = fixFileName + "无法删除重复数据，sheet格式问题，复制数据到新表" + ex.Message;
                    errorList.Add((errorExcel, errorExcelLog, fixFileName));
                }

            var endRowSource = targetSheet.Dimension.End.Row;

            targetSheet.InsertRow(endRowSource + 1, sourceDataList.Count);
            var colCount = targetSheet.Dimension.Columns;
            //多语言表不需要复制全部列
            if (fixFileName == "Localizations.xlsx")
            {
                colCount = 7;
            }
            var cellSource = targetSheet.Cells[endRowSource, 1, endRowSource, colCount];
            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var cellTarget = targetSheet.Cells[
                    endRowSource + 1 + m,
                    1,
                    endRowSource + 1 + m,
                    colCount
                ];
                cellSource.Copy(
                    cellTarget,
                    ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting
                        | ExcelRangeCopyOptionFlags.ExcludeMergedCells
                );
                cellSource.CopyStyles(cellTarget);
            }

            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var errorCell = sourceDataList[m][0];
                if (errorCell == null)
                {
                    errorExcelLog = sourceSheet.Name + "#表格尾行有，多余格式，清除";
                    MessageBox.Show(errorExcelLog);
                    return errorList;
                }

                var sourceCount = 0;
                foreach (var source in sourceKeyList)
                {
                    var cellCol = PubMetToExcel.FindSourceCol(
                        targetSheet,
                        2,
                        fixKeyList[sourceCount]
                    );
                    if (cellCol == -1)
                    {
                        if (fixKeyList[sourceCount] == "bgType")
                        {
                            sourceCount++;
                            continue;
                        }

                        errorExcel = i * 2 + 2;
                        errorExcelLog = fixFileName + "#表格字段#[" + fixKeyList[sourceCount] + "]未找到";
                        errorList.Add((errorExcel, errorExcelLog, fixFileName));
                        sourceCount++;
                        continue;
                    }

                    var cellTarget = targetSheet.Cells[endRowSource + 1 + m, cellCol];
                    var newStr = "";
                    if (int.TryParse(source, out var e))
                    {
                        var realCol = "";
                        if (fixFileName == "GuideDialogGroup.xlsx")
                            realCol = "GroupID";
                        else if (fixFileName == "GuideDialogBranch.xlsx")
                            realCol = "BranchID";
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(realCol)];
#pragma warning disable CS0252
                        if (sourceValue == "" || sourceValue == null)
                            continue;
#pragma warning restore CS0252
                        var str = sourceValue.ToString();
                        var digit = Math.Pow(10, e);
                        var repeatCount = 0;
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf(realCol)];
#pragma warning disable CS0252
                            if (repeatValue == "" || repeatValue == null)
                                continue;
#pragma warning restore CS0252
                            if (repeatValue == sourceValue)
                            {
#pragma warning disable CA1305
                                var newNum = long.Parse(str) * digit + repeatCount + 1;
#pragma warning restore CA1305
                                newStr = newStr + newNum + ",";
                                repeatCount++;
                            }
                        }

                        newStr = "[" + newStr.Substring(0, newStr.Length - 1) + "]";
                        cellTarget.Value = newStr;
                    }
                    else if (source == "枚举1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var scCol = classTitle.IndexOf(source);
                        var newId = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newId = classDataList[k][scCol + 1]?.ToString();
                                if(newId == null)
                                {
                                    MessageBox.Show($"{classDataList[k][0]}没有找到对应的ID值");
                                    return null;
                                }
                                break;
                            }
                        }

                        var sourceStr = cellTarget.Value?.ToString();
                        var reg = "\\d+";
                        if (string.IsNullOrEmpty(sourceStr))
                            continue;
                        var matches = Regex.Matches(sourceStr, reg);

                        var oldId = matches[0].Value;
                        if (newId != "")
                            sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色表情")
                    {
                        var sourceValue = sourceDataList[m]
                            [sourceTitle.IndexOf(source)]
                            ?.ToString();
                        for (var k = 0; k < emoDataList.Count; k++)
                        {
                            var targetValue = emoDataList[k][0].ToString();
                            if (targetValue == sourceValue)
                            {
                                var emoId = emoDataList[k][2];
                                if (emoId == null)
                                {
                                    emoId = "idle";
                                }
                                cellTarget.Value = emoId;
                                break;
                            }
                        }
                    }
                    else if (source == "触发分支")
                    {
                        var sourceValue = sourceDataList[m]
                            [sourceTitle.IndexOf(source)]
                            ?.ToString();
                        if (sourceValue == null || sourceValue == "" || sourceValue == "0")
                        {
                            sourceCount++;
                            continue;
                        }

                        var uniqueValues1 = new HashSet<string>();
                        var strBranch = "";
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k]
                                [sourceTitle.IndexOf("分支归属")]
                                ?.ToString();
                            if (repeatValue == null || repeatValue == "")
                                continue;
                            if (repeatValue == sourceValue)
                            {
                                var branchId = sourceDataList[k][sourceTitle.IndexOf("BranchID")];
                                if (uniqueValues1.Add((string)branchId))
                                    strBranch = strBranch + branchId + ",";
                            }
                        }

                        strBranch = "[" + strBranch.Substring(0, strBranch.Length - 1) + "]";
                        cellTarget.Value = strBranch;
                    }
                    else if (source == "分支多语言")
                    {
                        var newId = sourceDataList[m][sourceTitle.IndexOf("BranchID")]?.ToString();
                        var sourceStr = cellTarget.Value?.ToString();
                        if (string.IsNullOrEmpty(sourceStr))
                            continue;
                        var reg = "\\d+";
                        var matches = Regex.Matches(sourceStr, reg);
                        var oldId = matches[0].Value;
                        if (newId != "")
                            sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色换装1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m]
                            [sourceTitle.IndexOf("角色换装")]
                            ?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue =
                                    sourceValue2 == "1"
                                        ? (string)classDataList[k][scCol + 2].ToString()
                                        : "[]";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "角色换装2")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m]
                            [sourceTitle.IndexOf("角色换装")]
                            ?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue =
                                    sourceValue2 != "1"
                                        ? (string)classDataList[k][scCol + 3].ToString()
                                        : "";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "UI对话框")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("UI对话框")];
                        sourceValue = sourceValue == null ? "1" : sourceValue.ToString();
                        if (fixKeyList[sourceCount] == "bgType")
                            cellTarget.Value = sourceValue;
                    }
                    else
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)];
                        cellTarget.Value = sourceValue;
                    }

                    sourceCount++;
                }
            }

            if (errorExcel != 0)
                continue;
            var startRow = endRowSource + 1;
            int endRow2 = startRow + sourceDataList.Count - 1;
            if (
                fixFileName == "GuideDialogBranch.xlsx"
                || fixFileName == "GuideDialogGroup.xlsx"
                || fixFileName == "GuideDialogDetailSpecialSetting.xlsx"
                || fixFileName == "Localizations.xlsx"
            )
            {
                var uniqueValues = new HashSet<string>();
                for (var row = 4; row <= endRow2; row++)
                {
                    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";

                    if (uniqueValues.Contains(cellValue) || string.IsNullOrWhiteSpace(cellValue))
                    {
                        targetSheet.DeleteRow(row);
                        row--;
                        endRow2--;
                    }
                    else
                    {
                        uniqueValues.Add(cellValue);
                    }
                }
            }

            targetExcel.Save();
            targetExcel.Dispose();
            var excelCount = i / 2 + 1;
            app.StatusBar =
                "写入数据" + "<" + excelCount + "/" + fixDataList.Count / 2 + ">" + fixFileName;
        }

        return errorList;
    }

    //重构的多语言对话写入功能
    public static void AutoInsertDataByUdNew(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var excelPath = workBook.Path;

        // 获取基础数据
        var sourceSheet = workBook.Worksheets["多语言对话【模板】"];
        var sourceData = PubMetToExcel.ExcelDataToListBySelfToEnd(sourceSheet, 0, 1, 1);
        var sourceTitle = sourceData.Item1;
        List<List<object>> sourceDataList = sourceData.Item2;

        // 获取【数据修改】名称表
        var fixSheet = workBook.Worksheets["数据修改"];
        var fixSheetListObjects = fixSheet.ListObjects;
        var fixSheetValueAll = new Dictionary<string, Dictionary<(object, object), string>>();

        foreach (ListObject list in fixSheetListObjects)
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
            fixSheetValueAll[modelName] = modelValue;
        }

        // 获取【角色数据】名称表
        var roleSheet = workBook.Worksheets["角色数据"];
        var roleSheetListObjects = roleSheet.ListObjects;
        var roleSheetValueAll = new Dictionary<string, Dictionary<(object, object), string>>();

        foreach (ListObject list in roleSheetListObjects)
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
            roleSheetValueAll[modelName] = modelValue;
        }

        ErrorLogCtp.DisposeCtp();

        var errorExcelList = new List<List<(int, string, string)>>();
        if (errorExcelList == null)
            throw new ArgumentNullException(nameof(errorExcelList));

        string error = LanguageDialogDataByUdNew(
            sourceTitle,
            sourceDataList,
            fixSheetValueAll,
            roleSheetValueAll,
            excelPath
        );
        if (error != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(error);
        }

        NumDesAddIn.App.StatusBar = "导出完成";
        Marshal.ReleaseComObject(sourceSheet);
        Marshal.ReleaseComObject(fixSheet);
        Marshal.ReleaseComObject(roleSheet);
        Marshal.ReleaseComObject(workBook);
    }

    private static string LanguageDialogDataByUdNew(
        dynamic sourceTitle,
        List<List<object>> sourceDataList,
        dynamic fixSheetValueAll,
        dynamic roleSheetValueAll,
        string excelPath
    )
    {
        // 替换通配符生成数据
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var dicValue = new Dictionary<(string, string), List<string>>();

        string error = String.Empty;

        foreach (var fixSheet in fixSheetValueAll)
        {
            string fixSheetName = fixSheet.Key;
            string id = fixSheetName switch
            {
                "GuideDialogGroup.xlsx" => "GroupID",
                "GuideDialogDetail.xlsx" => "DetailID",
                "Localizations.xlsx" => "多语言KEY",
                "GuideDialogBranch.xlsx" => "BranchID",
                _ => ""
            };

            // 获取列索引
            int idIndex = sourceTitle.IndexOf(id);

            // 检查索引是否有效
            if (idIndex < 0)
            {
                error += $"{fixSheetName} 列名 '{id}' 在 多语言对话 中不存在\n";
            }

            // 提取列数据
            var idList = sourceDataList
                .Select(row => row != null && idIndex < row.Count ? row[idIndex] : null)
                .ToList();

            var rowData = new Dictionary<string, Dictionary<string, object>>();

            for (int idCount = 0; idCount < idList.Count; idCount++)
            {
                string itemId = idList[idCount]?.ToString() ?? "";
                if (itemId == "")
                {
                    continue;
                }
                var colData = new Dictionary<string, object>();
                foreach (var fixData in fixSheet.Value)
                {
                    string fixMethod = fixData.Value ?? "";
                    string fixKey = fixData.Key.Item2;

                    // 根据方法获得 fix 值
                    var fixValue = FixValueAnalysis(
                        idCount,
                        fixMethod,
                        sourceTitle,
                        sourceDataList,
                        dicValue,
                        roleSheetValueAll
                    );
                    if (fixValue.ToString().Contains("Error"))
                    {
                        error += fixValue + "\n";
                        LogDisplay.RecordLine(
                            "[{0}] , {1}【角色数据】中不存在",
                            DateTime.Now.ToString(CultureInfo.InvariantCulture),
                            fixValue
                        );
                    }
                    // 检查 fixValue 是否为空，避免覆盖已有数据
                    if (fixValue != null && !string.IsNullOrEmpty(fixValue.ToString()))
                    {
                        colData[fixKey] = fixValue;
                    }
                    else if (rowData.ContainsKey(itemId) && rowData[itemId] is { } existingColData)
                    {
                        // 如果 fixValue 为空，保留 rowData 中已有的值
                        if (existingColData.TryGetValue(fixKey, out var values))
                        {
                            colData[fixKey] = values;
                        }
                    }
                }

                // 更新 rowData
                rowData[itemId] = colData;
            }

            // 写入数据
            PubMetToExcel.SetExcelObjectEpPlus(
                excelPath,
                fixSheetName,
                out ExcelWorksheet targetSheet,
                out ExcelPackage targetExcel
            );
            
            //去重更新
            var newIdList = idList.Distinct().ToList();
            var rowsToDelete = new List<int>();
            foreach (var newId in newIdList)
            {
                var reDd = PubMetToExcel.FindSourceRow(targetSheet, 2, newId.ToString());
                if (reDd != -1)
                    rowsToDelete.Add(reDd);
            }

            rowsToDelete.Sort();
            rowsToDelete.Reverse();

            foreach (var rowToDelete in rowsToDelete)
                try
                {
                    targetSheet.DeleteRow(rowToDelete, 1);
                }
                catch (Exception)
                {
                    LogDisplay.RecordLine(
                        "[{0}] , {1}",
                        DateTime.Now.ToString(CultureInfo.InvariantCulture),
                        "sheet表有问题无法删除"
                    );
                }

            var writeCol = targetSheet.Dimension.End.Column;
            bool dataWritten = false; // 标志是否有实际写入
            var dataRepeatWritten = new HashSet<string>();
            foreach (var row in rowData)
            {
                string itemId = row.Key;
                if (itemId == "")
                    continue;
                var writeRow = targetSheet.Dimension.End.Row + 1;

                HashSet<string> processedKeys = new HashSet<string>();
                for (int j = 2; j <= writeCol; j++)
                {
                    var cellTitle = targetSheet.Cells[2, j].Value?.ToString() ?? "";

                    if (cellTitle == "")
                        continue;

                    // 使用 LINQ 查询判断字典中是否包含指定的值
                    var matchingKey = row.Value.Keys.FirstOrDefault(key => key.Equals(cellTitle));
                    var isContains = processedKeys.Contains(cellTitle);
                    if (matchingKey != null && !isContains)
                    {
                        processedKeys.Add(cellTitle);
                        var cellRealValue = row.Value[cellTitle];
                        // 空ID判断
                        if (j == 2 && (string)cellRealValue == string.Empty)
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
                            dataRepeatWritten.Add(cellRealValue?.ToString());
                        }

                        // 实际写入
                        var cell = targetSheet.Cells[writeRow, j];
                        cell.Value = cellRealValue;
                        dataWritten = true;
                    }
                }
            }
            if (dataWritten) // 只有在写入数据时才保存
            {
                targetExcel.Save();
                NumDesAddIn.App.StatusBar = $"导出：{fixSheetName}";
            }
            targetSheet.Dispose();
        }
        return error;
    }

    private static string FixValueAnalysis(
        int idCount,
        string fixMethod,
        dynamic sourceTitle,
        dynamic sourceDataList,
        Dictionary<(string, string), List<string>> dicValue,
        dynamic roleSheetValueAll
    )
    {
        string cellRealValue = fixMethod;
        string wildcardPattern = "#(.*?)#";
        string wildcardValuePattern = "-";

        MatchCollection matches = Regex.Matches(fixMethod, wildcardPattern);

        foreach (Match match in matches)
        {
            var wildcard = match.Groups[1].Value;

            var wildcardValueSplit = Regex.Split(wildcard, wildcardValuePattern);
            string funName = wildcardValueSplit.ElementAtOrDefault(0) ?? "";
            string funDy1 = wildcardValueSplit.ElementAtOrDefault(1) ?? "";
            string funDy2 = wildcardValueSplit.ElementAtOrDefault(2) ?? "";
            string funDy3 = wildcardValueSplit.ElementAtOrDefault(3) ?? "";

            string fixWildcardValue = funName switch
            {
                "Dic" => Dic(funDy1, funDy2, funDy3),
                "Find" => Find(funDy1, funDy2, funDy3),
                "Merge" => Merge(funDy1, funDy2, funDy3),
                _ => GetValue(funName)
            };
            fixWildcardValue ??= "";
            if (fixWildcardValue.Contains("Error"))
            {
                return fixWildcardValue;
            }
            cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
        }
        return cellRealValue;

        string Dic(string funDy1, string funDy2, string funDy3)
        {
            var itemValue = sourceDataList[idCount][sourceTitle.IndexOf(funDy1)];
            itemValue = itemValue != null ? itemValue.ToString() : string.Empty;

            if (!dicValue.ContainsKey((funDy2, itemValue)))
            {
                dicValue[(funDy2, itemValue)] = new List<string>();
            }

            var rawValue = sourceDataList[idCount][sourceTitle.IndexOf(funDy2)];
            string value = rawValue != null ? rawValue.ToString() : string.Empty;
            dicValue[(funDy2, itemValue)].Add(value);

            if (funDy3 != "0")
            {
                dicValue[(funDy2, itemValue)] = dicValue[(funDy2, itemValue)]
                    .Where(new Func<string, bool>(values => !string.IsNullOrEmpty(values))) // 过滤掉 null 和空字符串
                    .Distinct()
                    .ToList();
            }
            // list变字符串
            string result = string.Join(",", dicValue[(funDy2, itemValue)]);
            return result;
        }

        string Find(string funDy1, string funDy2, string funDy3)
        {
            var findSheet = roleSheetValueAll[funDy1];
            var findValue = sourceDataList[idCount][sourceTitle.IndexOf(funDy2)];
            if (findValue == null)
            {
                return String.Empty;
            }

            string result;
            try
            {
                result = findSheet[((object)findValue, (object)funDy3)];
            }
            catch
            {
                LogDisplay.RecordLine(
                    "[{0}] , {1}【角色数据】中不存在",
                    DateTime.Now.ToString(CultureInfo.InvariantCulture),
                    findValue
                );
                result = $"Error#{findValue}#在【角色数据】中不存在";
            }
            return result;
        }
        string Merge(string funDy1, string funDy2, string funDy3)
        {
            string result = String.Empty;

            string itemValue1;
            string itemValue2;
            string itemValue3;

            if (sourceTitle.IndexOf(funDy1) == -1)
            {
                itemValue1 = funDy1;
            }
            else
            {
                itemValue1 = sourceDataList[idCount][sourceTitle.IndexOf(funDy1)]?.ToString();
            }
            if (sourceTitle.IndexOf(funDy2) == -1)
            {
                itemValue2 = funDy2;
            }
            else
            {
                itemValue2 = sourceDataList[idCount][sourceTitle.IndexOf(funDy2)]?.ToString();
            }
            if (sourceTitle.IndexOf(funDy3) == -1)
            {
                itemValue3 = funDy3;
            }
            else
            {
                itemValue3 = sourceDataList[idCount][sourceTitle.IndexOf(funDy3)]?.ToString();
            }

            if (itemValue3 != null)
            {
                result = $"{itemValue1}{itemValue2}{itemValue3}";
            }
            return result;
        }

        string GetValue(string funName)
        {
            var getValueCol = sourceTitle.IndexOf(funName);
            var getValue = sourceDataList[idCount][getValueCol];
            var result = getValue?.ToString();
            return result;
        }
    }
}

//模版数据写入-老方法，需要填写大量字段修改参数
public static class ExcelDataAutoInsertMulti
{
    public static void InsertData(dynamic isMulti)
    {
        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var colsCount = sheet.UsedRange.Columns.Count;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdCol = title.IndexOf("初始模板");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        var fixKeyCol = title.IndexOf("修改字段");
        var baseIdCol = title.IndexOf("模板期号");
        var creatIdCol = title.IndexOf("创建期号");
        var commentValue = data[2][baseIdCol];
        var writeMode = data[2][creatIdCol];
        ErrorLogCtp.DisposeCtp();
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        var addValue = (int)data[0][creatIdCol] - (int)data[0][baseIdCol];
        var rowCount = 2;
        var colFixKeyCount = colsCount - fixKeyCol;
        var modelId = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdCol, rowCount);
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(
            data,
            sheetNameCol,
            modelIdNewCol,
            rowCount
        );
        var fixKey = PubMetToExcel.ExcelDataToDictionary(
            data,
            sheetNameCol,
            fixKeyCol,
            rowCount,
            colFixKeyCount
        );
        var ignoreExcel = PubMetToExcel.ExcelDataToDictionary(
            data,
            sheetNameCol,
            creatIdCol,
            rowCount
        );
        var errorExcelList = new List<List<(string, string, string)>>();
        var excelCount = 1;
        foreach (var key in modelId)
        {
            var excelName = key.Key;
            var ignore = ignoreExcel[excelName][0].Item1[0, 0];
            if (ignore != null)
            {
                var ignoreStr = ignore.ToString();
                if (ignoreStr == "跳过")
                {
                    NumDesAddIn.App.StatusBar = "跳过" + "<" + excelName;
                    excelCount++;
                    continue;
                }
            }

            List<(string, string, string)> error = ExcelDataWrite(
                modelId,
                modelIdNew,
                fixKey,
                excelPath,
                excelName,
                addValue,
                isMulti,
                commentValue,
                cellColor,
                writeMode
            );
            NumDesAddIn.App.StatusBar =
                "写入数据" + "<" + excelCount + "/" + modelId.Count + ">" + excelName;
            errorExcelList.Add(error);
            excelCount++;
        }

        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            NumDesAddIn.App.StatusBar = "完成写入";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static void RightClickInsertData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        NumDesAddIn.App.StatusBar = false;
        var sw = new Stopwatch();
        sw.Start();

        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var colsCount = sheet.UsedRange.Columns.Count;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdCol = title.IndexOf("初始模板");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        var fixKeyCol = title.IndexOf("修改字段");
        var baseIdCol = title.IndexOf("模板期号");
        var creatIdCol = title.IndexOf("创建期号");
        var commentValue = data[2][baseIdCol];
        var writeMode = data[2][creatIdCol];
        ErrorLogCtp.DisposeCtp();
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        var addValue = (int)data[0][creatIdCol] - (int)data[0][baseIdCol];
        var rowCount = 2;
        var colFixKeyCount = colsCount - fixKeyCol;
        var modelId = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdCol, rowCount);
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(
            data,
            sheetNameCol,
            modelIdNewCol,
            rowCount
        );
        var fixKey = PubMetToExcel.ExcelDataToDictionary(
            data,
            sheetNameCol,
            fixKeyCol,
            rowCount,
            colFixKeyCount
        );
        var errorExcelList = new List<List<(string, string, string)>>();
        var cell = NumDesAddIn.App.Selection;
        var rowStart = cell.Row;
        var rowCountNew = cell.Rows.Count;
        var rowEnd = rowStart + rowCountNew - 1;
        var excelList = new List<string>();
        for (int i = rowStart; i <= rowEnd; i++)
        {
            var excelName = data[i - 2][sheetNameCol];
            excelList.Add((string)excelName);
        }

        var newExcelList = excelList
            .Where(excelName => !string.IsNullOrEmpty(excelName))
            .Distinct()
            .ToList();
        for (var i = 0; i < newExcelList.Count; i++)
        {
            var excelName = newExcelList[i];
            if (excelName == null)
                continue;
            List<(string, string, string)> error = ExcelDataWrite(
                modelId,
                modelIdNew,
                fixKey,
                excelPath,
                excelName,
                addValue,
                false,
                commentValue,
                cellColor,
                writeMode
            );
            NumDesAddIn.App.StatusBar =
                "写入数据" + "<" + i + "/" + newExcelList.Count + ">" + excelName;
            errorExcelList.Add(error);
        }

        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            sw.Stop();
            var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
            NumDesAddIn.App.StatusBar = "完成写入，用时：" + ts2.ToString(CultureInfo.InvariantCulture);
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
        sw.Stop();
        var ts3 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        NumDesAddIn.App.StatusBar = "完成写入:有错误，用时：" + ts3.ToString(CultureInfo.InvariantCulture);
    }

    public static List<(string, string, string)> ExcelDataWrite(
        dynamic modelId,
        dynamic modelIdNew,
        dynamic fixKey,
        dynamic excelPath,
        dynamic excelName,
        dynamic addValue,
        dynamic modeThread,
        dynamic commentValue,
        dynamic cellBackColor,
        dynamic writeMode
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorExcelLog = "";
        var errorList = new List<(string, string, string)>();
        string path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        var sheetRealName = "Sheet1";
        string excelRealName = excelName;
        if (excelName.Contains("#"))
        {
            var excelRealNameGroup = excelName.Split("#");
            excelRealName = excelRealNameGroup[0];
            sheetRealName = excelRealNameGroup[1];
        }

        switch (excelRealName)
        {
            case "Localizations.xlsx":
                path = newPath + @"\Excels\Localizations\Localizations.xlsx";
                break;
            case "UIConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIConfigs.xlsx";
                break;
            case "UIItemConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIItemConfigs.xlsx";
                break;
            default:
                path = excelPath + @"\" + excelRealName;
                break;
        }

        var fileExists = File.Exists(path);
        if (fileExists == false)
        {
            errorExcelLog = excelRealName + "不存在表格文件";
            errorList.Add((excelRealName, errorExcelLog, excelRealName));
            return errorList;
        }

        var excel = new ExcelPackage(new FileInfo(path));
        ExcelWorkbook workBook;
        try
        {
            workBook = excel.Workbook;
        }
        catch (Exception ex)
        {
            errorExcelLog = excelRealName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelRealName, errorExcelLog, excelRealName));
            return errorList;
        }

        ExcelWorksheet sheet;
        try
        {
            sheet = workBook.Worksheets[sheetRealName];
        }
        catch (Exception ex)
        {
            errorExcelLog = excelRealName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelRealName, errorExcelLog, excelRealName));
            return errorList;
        }

        sheet ??= workBook.Worksheets[0];
        foreach (var cell in sheet.Cells)
            if (cell.Formula is { Length: > 0 })
            {
                errorList.Add((excelRealName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
                return errorList;
            }

        var writeIdList = ExcelDataWriteIdGroup(excelName, addValue, sheet, fixKey, modelId);
        PubMetToExcel.RepeatValue2(sheet, 4, 2, writeIdList.Item1);

        var colCount = sheet.Dimension.Columns;
        //多语言表不需要复制全部列
        if (excelRealName == "Localizations.xlsx")
        {
            colCount = 7;
        }
        writeIdList = ExcelDataWriteIdGroup(excelName, addValue, sheet, fixKey, modelId);
        var writeRow = writeIdList.Item2;
        if (writeRow == -1)
        {
            errorExcelLog = excelName + "#找不到" + writeIdList.Item1[0];
            errorList.Add((excelName, errorExcelLog, excelName));
            return errorList;
        }

        for (var excelMulti = 0; excelMulti < modelId[excelName].Count; excelMulti++)
        {
            var startValue = modelId[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelId[excelName][excelMulti].Item1[1, 0].ToString();

            var startRowSource = PubMetToExcel.FindSourceRow(sheet, 2, startValue);
            if (startRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
                errorList.Add((startValue, errorExcelLog, excelName));
                return errorList;
            }

            var endRowSource = PubMetToExcel.FindSourceRow(sheet, 2, endValue);
            if (endRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]未找到(序号出错)";
                errorList.Add((endValue, errorExcelLog, excelName));
                return errorList;
            }

            if (endRowSource - startRowSource < 0)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]起始、终结ID顺序反了";
                errorList.Add((endValue, errorExcelLog, excelName));
                return errorList;
            }
            //复制数据
            if (excelRealName.Contains("Recharge"))
            {
                writeRow = sheet.Dimension.End.Row;
            }
            var count = endRowSource - startRowSource + 1;
            sheet.InsertRow(writeRow + 1, count);
            var cellSource = sheet.Cells[startRowSource, 1, endRowSource, colCount];
            var cellTarget = sheet.Cells[writeRow + 1, 1, writeRow + count, colCount];
            cellTarget.Value = cellSource.Value;
            cellTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cellTarget.Style.Fill.BackgroundColor.SetColor(cellBackColor);

            cellTarget.Style.Font.Name = "微软雅黑";
            cellTarget.Style.Font.Size = 10;
            cellTarget.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            //修改数据
            var fixItem = fixKey[excelName][excelMulti].Item1;
            errorList = modeThread
                ? (List<(string, string, string)>)
                    MultiWrite(
                        excelName,
                        addValue,
                        fixItem,
                        sheet,
                        count,
                        startRowSource,
                        errorList,
                        commentValue,
                        writeRow
                    )
                : (List<(string, string, string)>)
                    SingleWrite(
                        excelName,
                        addValue,
                        fixItem,
                        sheet,
                        count,
                        startRowSource,
                        errorList,
                        commentValue,
                        writeRow
                    );
            writeRow += count;
        }

        excel.Save();
        excel.Dispose();
        errorList.Add(("-1", errorExcelLog, excelName));
        return errorList;
    }

    private static List<(string, string, string)> SingleWrite(
        dynamic excelName,
        dynamic addValue,
        dynamic fixItem,
        ExcelWorksheet sheet,
        dynamic count,
        dynamic startRowSource,
        List<(string, string, string)> errorList,
        dynamic commentValue,
        int writeRow
    )
    {
        for (var colMulti = 0; colMulti < fixItem.GetLength(1); colMulti++)
        {
            string excelKey = fixItem[0, colMulti];
            //遇到excelKey这类自定义写入的，继续按规则写入，否则进入批量替换，不更改的字段需要写1#0；
            if (excelKey == null)
                continue;
            var excelFileFixKey = PubMetToExcel.FindSourceCol(sheet, 2, excelKey);
            if (excelFileFixKey == -1)
            {
                var errorExcelLog = excelName + "#【初始模板】#[" + excelKey + "]未找到(字段出错)";
                errorList.Add((excelKey, errorExcelLog, excelName));
                continue;
            }

            string excelKeyMethod = fixItem[1, colMulti]?.ToString();
            for (var i = 0; i < count; i++)
            {
                var cellSource = sheet.Cells[startRowSource + i, excelFileFixKey];
                var rowId = sheet.Cells[startRowSource + i, 2];
                var cellCol = sheet.Cells[2, excelFileFixKey].Value?.ToString();
                var cellFix = sheet.Cells[writeRow + 1 + i, excelFileFixKey];

                if (cellCol != null && cellCol.Contains("#") && commentValue != null)
                {
                    string[] baseParts = commentValue.Split("#");
                    var cellValue = cellFix.Value?.ToString();

                    int partCount = 0;
                    foreach (var item in baseParts)
                    {
                        var parts = item.Split("-");
                        var replaceValue = parts[0];
                        //备注的全量替换
                        if (replaceValue.Contains("***"))
                        {
                            replaceValue = baseParts[partCount].Replace("***", "");
                            cellValue = replaceValue;
                        }
                        else
                        {
                            var pattern = parts[1];
                            if (cellValue != null)
                                cellValue = Regex.Replace(cellValue, pattern, replaceValue);
                        }

                        partCount++;
                    }

                    cellFix.Value = cellValue;
                }
                else
                {
                    string cellFixValue;
                    //固定值
                    string baseValue = excelKeyMethod ?? "";
                    if (baseValue.Contains("***"))
                    {
                        baseValue = baseValue.Replace("***", "");
                        cellFixValue = baseValue;
                    }
                    //自增值
                    else
                    {
                        if (cellSource.Value == null)
                            continue;

                        if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0")
                            continue;

                        var fixValueList = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                        cellFixValue = ExcelDataAutoInsert.StringRegPlace(
                            cellSource.Value.ToString(),
                            fixValueList,
                            addValue
                        );
                    }
                    if (cellFixValue == "^error^")
                    {
                        string errorExcelLog =
                            excelName + "#" + rowId.Value + "#【修改模式】#[" + excelKey + "]字段方法写错";
                        errorList.Add((excelKey, errorExcelLog, excelName));
                    }

                    cellFix.Value = double.TryParse(cellFixValue, out double number)
                        ? number
                        : cellFixValue;
                }
            }
        }

        return errorList;
    }

    private static List<(string, string, string)> MultiWrite(
        dynamic excelName,
        dynamic addValue,
        dynamic fixItem,
        ExcelWorksheet sheet,
        dynamic count,
        dynamic startRowSource,
        List<(string, string, string)> errorList,
        dynamic commentValue,
        int writeRow
    )
    {
        var colCoinMulti = fixItem.GetLength(1);
        var colThreadCount = 8;
        int colBatchSize = colCoinMulti / colThreadCount;
        Parallel.For(
            0,
            colThreadCount,
            e =>
            {
                var startRow = e * colBatchSize;
                var endRow = (e + 1) * colBatchSize;
                if (e == colThreadCount - 1)
                    endRow = colCoinMulti;
                for (var k = startRow; k < endRow; k++)
                {
                    string excelKey = fixItem[0, k];
                    if (excelKey == null)
                        continue;
                    var excelFileFixKey = PubMetToExcel.FindSourceCol(sheet, 2, excelKey);
                    if (excelFileFixKey == -1)
                    {
                        var errorExcelLog = excelName + "#【初始模板】#[" + excelKey + "]未找到(字段出错)";
                        errorList.Add((excelKey, errorExcelLog, excelName));
                        continue;
                    }

                    string excelKeyMethod = fixItem[1, k]?.ToString();

                    var rowThreadCount = 4;
                    int rowBatchSize = count / rowThreadCount;
                    Parallel.For(
                        0,
                        rowThreadCount,
                        i =>
                        {
                            var startCol = i * rowBatchSize;
                            var endCol = (i + 1) * rowBatchSize;
                            if (i == rowThreadCount - 1)
                                endCol = count;

                            for (var j = startCol; j < endCol; j++)
                            {
                                var cellSource = sheet.Cells[startRowSource + j, excelFileFixKey];
                                var cellCol = sheet.Cells[2, excelFileFixKey].Value?.ToString();
                                var cellFix = sheet.Cells[writeRow + j + 1, excelFileFixKey];
                                var rowId = sheet.Cells[startRowSource + j, 2];

                                if (
                                    cellCol != null
                                    && cellCol.Contains("#")
                                    && commentValue != null
                                )
                                {
                                    string[] baseParts = commentValue.Split("#");
                                    var cellValue = cellFix.Value?.ToString();

                                    int partCount = 0;
                                    foreach (var item in baseParts)
                                    {
                                        var parts = item.Split("-");
                                        var replaceValue = parts[0];
                                        //备注的全量替换
                                        if (replaceValue.Contains("***"))
                                        {
                                            replaceValue = baseParts[partCount].Replace("***", "");
                                            cellValue = replaceValue;
                                        }
                                        else
                                        {
                                            var pattern = parts[1];
                                            if (cellValue != null)
                                                cellValue = Regex.Replace(cellValue, pattern, replaceValue);
                                        }

                                        partCount++;
                                    }

                                    cellFix.Value = cellValue;
                                }
                                else
                                {
                                    string cellFixValue;
                                    //固定值
                                    string baseValue = excelKeyMethod ?? "";
                                    if (baseValue.Contains("***"))
                                    {
                                        baseValue = baseValue.Replace("***", "");
                                        cellFixValue = baseValue;
                                    }
                                    //自增值
                                    else
                                    {
                                        if (cellSource.Value == null)
                                            continue;

                                        if (
                                            cellSource.Value.ToString() == ""
                                            || cellSource.Value.ToString() == "0"
                                        )
                                            continue;
                                        var fixValueList = ExcelDataAutoInsert.CellFixValueKeyList(
                                            excelKeyMethod
                                        );
                                        cellFixValue = ExcelDataAutoInsert.StringRegPlace(
                                            cellSource.Value.ToString(),
                                            fixValueList,
                                            addValue
                                        );
                                    }
                                    if (cellFixValue == "^error^")
                                    {
                                        string errorExcelLog =
                                            excelName
                                            + "#"
                                            + rowId.Value
                                            + "#【修改模式】#["
                                            + excelKey
                                            + "]字段方法写错";
                                        errorList.Add((excelKey, errorExcelLog, excelName));
                                    }

                                    cellFix.Value = double.TryParse(cellFixValue, out double number)
                                        ? number
                                        : cellFixValue;
                                }
                            }
                        }
                    );
                }
            }
        );
        return errorList;
    }

    public static (List<string>, int) ExcelDataWriteIdGroup(
        dynamic excelName,
        dynamic addValue,
        ExcelWorksheet sheet,
        dynamic fixKey,
        dynamic modelId
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var excelFileFixKey = 2;
        var writeIdList = new List<string>();
        var lastRow = 0;
        for (var excelMulti = 0; excelMulti < modelId[excelName].Count; excelMulti++)
        {
            var startValue = modelId[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelId[excelName][excelMulti].Item1[1, 0].ToString();
            var startRowSource = PubMetToExcel.FindSourceRow(sheet, 2, startValue);
            var endRowSource = PubMetToExcel.FindSourceRow(sheet, 2, endValue);
            if (startRowSource == -1 || endRowSource == -1)
            {
                var writeIdList2 = new List<string> { startValue + "#" + endValue };
                return (writeIdList2, -1);
            }

            string excelKeyMethod = fixKey[excelName][excelMulti].Item1[1, 0]?.ToString();
            var count = endRowSource - startRowSource + 1;
            for (var i = 0; i < count; i++)
            {
                var cellSource = sheet.Cells[startRowSource + i, excelFileFixKey];

                string cellFixValue;
                //固定值
                string baseValue = excelKeyMethod ?? "";
                if (baseValue.Contains("***"))
                {
                    baseValue = baseValue.Replace("***", "");
                    cellFixValue = baseValue;
                }
                else
                {
                    if (cellSource.Value == null)
                        continue;
                    if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0")
                        continue;

                    var temp1 = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                    cellFixValue = ExcelDataAutoInsert.StringRegPlace(
                        cellSource.Value.ToString(),
                        temp1,
                        addValue
                    );
                }

                writeIdList.Add(cellFixValue);
            }

            if (lastRow < endRowSource)
                lastRow = endRowSource;
        }

        return (writeIdList, lastRow);
    }
}

//模版数据写入-新方法，只需要填写少量字段修改参数（包含不更改），其他数据进行自动匹配替换关键字
public static class ExcelDataAutoInsertMultiNew
{
    private static dynamic _indexWk;
    private static dynamic _sheet;
    private static dynamic _excelPath;
    private static dynamic _sheetData;
    private static dynamic _title;
    private static dynamic _data;
    private static dynamic _sheetNameCol;
    private static dynamic _modelIdCol;
    private static dynamic _modelIdNewCol;
    private static dynamic _fixKeyCol;
    private static dynamic _baseIdCol;
    private static dynamic _creatIdCol;
    private static dynamic _baseCommentCol;
    private static dynamic _creatCommentCol;
    private static dynamic _specialReplaceValueCol;
    private static dynamic _replaceValues;
    private static dynamic _colorCell;
    private static Color _cellColor;
    private static dynamic _addValue;
    private static dynamic _rowCount;
    private static dynamic _colFixKeyCount;
    private static dynamic _modelId;
    private static dynamic _modelIdNew;
    private static dynamic _fixKey;
    private static dynamic _ignoreExcel;
    private static dynamic _commentValue;
    private static dynamic _specialReplaceValue;
    private static dynamic _errorExcelList;

    //初始化参数
    private static void InitializeVariables()
    {
        _indexWk = NumDesAddIn.App.ActiveWorkbook;
        _sheet = NumDesAddIn.App.ActiveSheet;
        _excelPath = _indexWk.Path;

        _sheetData = PubMetToExcel.ExcelDataToList(_sheet);
        _title = _sheetData.Item1;
        _data = _sheetData.Item2;
        _sheetNameCol = _title.IndexOf("表名");
        _modelIdCol = _title.IndexOf("初始模板");
        _modelIdNewCol = _title.IndexOf("实际模板(上一期)");
        _fixKeyCol = _title.IndexOf("修改字段");
        _baseIdCol = _title.IndexOf("模板期号");
        _creatIdCol = _title.IndexOf("创建期号");
        _baseCommentCol = _title.IndexOf("初始备注");
        _creatCommentCol = _title.IndexOf("当前备注");
        _specialReplaceValueCol = _title.IndexOf("专属替换");
        _replaceValues = _data[2][_baseIdCol];

        //记录日志
        LogDisplay.RecordLine("[{0}] , 【表名】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【初始模板】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【实际模板(上一期)】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【修改字段】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【模板期号】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【创建期号】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【初始备注】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【当前备注】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);
        LogDisplay.RecordLine("[{0}] , 【专属替换】所在列：{1}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            _sheetNameCol);



        _colorCell = _sheet.Cells[6, 1];
        _cellColor = PubMetToExcel.GetCellBackgroundColor(_colorCell);
        _addValue = (int)_data[0][_creatIdCol] - (int)_data[0][_baseIdCol];
        _rowCount = 2;
        _colFixKeyCount = _baseCommentCol - _fixKeyCol;
        _modelId = PubMetToExcel.ExcelDataToDictionary(
            _data,
            _sheetNameCol,
            _modelIdCol,
            _rowCount
        );
        _modelIdNew = PubMetToExcel.ExcelDataToDictionary(
            _data,
            _sheetNameCol,
            _modelIdNewCol,
            _rowCount
        );
        _fixKey = PubMetToExcel.ExcelDataToDictionary(
            _data,
            _sheetNameCol,
            _fixKeyCol,
            _rowCount,
            _colFixKeyCount
        );
        _ignoreExcel = PubMetToExcel.ExcelDataToDictionary(
            _data,
            _sheetNameCol,
            _creatIdCol,
            _rowCount
        );
        _commentValue = PubMetToExcel.ExcelDataToDictionary(
            _data,
            _baseCommentCol,
            _creatCommentCol,
            1
        );
        _specialReplaceValue = PubMetToExcel.ExcelDataToDictionary(
            _data,
            _sheetNameCol,
            _specialReplaceValueCol,
            _rowCount
        );
        _errorExcelList = new List<List<(string, string, string)>>();
    }

    public static void InsertDataNew(dynamic isMulti)
    {
        InitializeVariables();
        ErrorLogCtp.DisposeCtp();
        var excelCount = 1;
        foreach (var key in _modelId)
        {
            var excelName = key.Key;
            var ignore = _ignoreExcel[excelName][0].Item1[0, 0];
            if (ignore != null)
            {
                var ignoreStr = ignore.ToString();
                if (ignoreStr == "跳过")
                {
                    NumDesAddIn.App.StatusBar = "跳过" + "<" + excelName;
                    excelCount++;
                    continue;
                }
            }

            List<(string, string, string)> error = CopyData(excelName);
            NumDesAddIn.App.StatusBar =
                "写入数据" + "<" + excelCount + "/" + _modelId.Count + ">" + excelName;
            _errorExcelList.Add(error);
            excelCount++;
        }

        var errorLog = PubMetToExcel.ErrorLogAnalysis(_errorExcelList, _sheet);
        if (errorLog == "")
        {
            NumDesAddIn.App.StatusBar = "完成写入";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static void RightClickInsertDataNew(CommandBarButton ctrl, ref bool cancelDefault)
    {
        InitializeVariables();
        NumDesAddIn.App.StatusBar = false;
        var sw = new Stopwatch();
        sw.Start();
        var cell = NumDesAddIn.App.Selection;
        var rowStart = cell.Row;
        var rowCountNew = cell.Rows.Count;
        var rowEnd = rowStart + rowCountNew - 1;
        var excelList = new List<string>();

        for (int i = rowStart; i <= rowEnd; i++)
        {
            var excelName = _data[i - 2][_sheetNameCol];
            excelList.Add((string)excelName);
        }

        var newExcelList = excelList
            .Where(excelName => !string.IsNullOrEmpty(excelName))
            .Distinct()
            .ToList();
        for (var i = 0; i < newExcelList.Count; i++)
        {
            var excelName = newExcelList[i];
            if (excelName == null)
                continue;
            List<(string, string, string)> error = CopyData(excelName);
            NumDesAddIn.App.StatusBar =
                "写入数据" + "<" + i + "/" + newExcelList.Count + ">" + excelName;
            _errorExcelList.Add(error);
        }

        var errorLog = PubMetToExcel.ErrorLogAnalysis(_errorExcelList, _sheet);
        if (errorLog == "")
        {
            sw.Stop();
            var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
            NumDesAddIn.App.StatusBar = "完成写入，用时：" + ts2.ToString(CultureInfo.InvariantCulture);
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
        sw.Stop();
        var ts3 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        NumDesAddIn.App.StatusBar = "完成写入:有错误，用时：" + ts3.ToString(CultureInfo.InvariantCulture);
    }

    public static List<(string, string, string)> CopyData(string excelName)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorExcelLog = "";
        List<(string, string, string)> errorList = PubMetToExcel.SetExcelObjectEpPlus(
            _excelPath,
            excelName,
            out ExcelWorksheet sheet,
            out ExcelPackage excel
        );
        if (excel == null)
        {
            LogDisplay.RecordLine(
                "[{0}] , {1}不存在，看看是否重命名了",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                excelName
            );
        }

        if (excel != null)
        {
            var excelRealName = excel.File.Name;

            foreach (var cell in sheet.Cells)
                if (cell.Formula is { Length: > 0 })
                {
                    errorList.Add(
                        ($"{excelRealName}#{sheet.Name}", @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@")
                    );
                    return errorList;
                }

            //查找是否已经写入过新ID，如果写入过，则删除
            var writeIdList = GetElementIdGroup(excelName, sheet, _modelIdNew, true);

            var writeRow = writeIdList.Item2;
            if (writeRow == -9527)
            {
                errorExcelLog = excelName + "#重复值#" + writeIdList.Item1[0];
                errorList.Add((excelName, errorExcelLog, excelName));
                return errorList;
            }

            //多语言表不需要复制全部列
            var colCount = sheet.Dimension.Columns;
            if (excelRealName == "Localizations.xlsx")
            {
                colCount = 7;
            }

            //获取老ID所在行列信息，准备复制
            writeIdList = GetElementIdGroup(excelName, sheet, _modelId);

            writeRow = writeIdList.Item2;
            if (writeRow == -1)
            {
                errorExcelLog = excelName + "#找不到" + writeIdList.Item1[0];
                errorList.Add((excelName, errorExcelLog, excelName));
                return errorList;
            }

            for (var excelMulti = 0; excelMulti < _modelId[excelName].Count; excelMulti++)
            {
                var startValue = _modelId[excelName][excelMulti].Item1[0, 0].ToString();
                var endValue = _modelId[excelName][excelMulti].Item1[1, 0].ToString();

                if (string.IsNullOrEmpty(startValue) || string.IsNullOrEmpty(endValue))
                {
                    errorExcelLog = excelName + "#【初始模板】#起始或结束值为空";
                    errorList.Add((startValue ?? "空值", errorExcelLog, excelName));
                    return errorList;
                }

                var startRowSource = PubMetToExcel.FindSourceRow(sheet, 2, startValue);
                if (startRowSource == -1)
                {
                    errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
                    errorList.Add((startValue, errorExcelLog, excelName));
                    return errorList;
                }

                var endRowSource = PubMetToExcel.FindSourceRow(sheet, 2, endValue);
                if (endRowSource == -1)
                {
                    errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]未找到(序号出错)";
                    errorList.Add((endValue, errorExcelLog, excelName));
                    return errorList;
                }

                if (endRowSource - startRowSource < 0)
                {
                    errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]起始、终结ID顺序反了";
                    errorList.Add((endValue, errorExcelLog, excelName));
                    return errorList;
                }

                //复制数据
                if (excelRealName.Contains("Recharge"))
                {
                    writeRow = sheet.Dimension.End.Row;
                }
                var count = endRowSource - startRowSource + 1;
                sheet.InsertRow(writeRow + 1, count);
                var cellSource = sheet.Cells[startRowSource, 1, endRowSource, colCount];
                var cellTarget = sheet.Cells[writeRow + 1, 1, writeRow + count, colCount];
                cellTarget.Value = cellSource.Value;
                cellTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cellTarget.Style.Fill.BackgroundColor.SetColor(_cellColor);

                cellTarget.Style.Font.Name = "微软雅黑";
                cellTarget.Style.Font.Size = 10;
                cellTarget.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                //修改数据
                var fixItem = _fixKey[excelName][excelMulti].Item1;
                //专属替换
                var specialValue = _specialReplaceValue[excelName][excelMulti].Item1;
                errorList = FixData(
                    excelName,
                    fixItem,
                    sheet,
                    count,
                    startRowSource,
                    writeRow,
                    errorList,
                    specialValue
                );
                writeRow += count;
            }
        }

        if (excel != null)
        {
            excel.Save();
            excel.Dispose();
        }

        errorList.Add(("-1", errorExcelLog, excelName));
        return errorList;
    }

    private static List<(string, string, string)> FixData(
        dynamic excelName,
        dynamic fixItem,
        ExcelWorksheet wkSheet,
        dynamic count,
        dynamic startRowSource,
        int writeRow,
        dynamic errorList,
        dynamic specialValue
    )
    {
        // 获取工作表的行数和列数
        var colCount = wkSheet.Dimension.Columns;

        //遍历目标表字段（区分自定义还是批量替换字段）
        for (var cellCol = 2; cellCol <= colCount; cellCol++)
        {
            var cellKey = wkSheet.Cells[2, cellCol].Value?.ToString() ?? "";
            var excelKeyFun = PubMetToExcel.FindValueInFirstRow(fixItem, cellKey);

            //通用修改（替换）遇到excelKey这类自定义写入的，继续按规则写入，否则进入批量替换，不更改的字段需要写1#0；
            if (cellKey != "" && excelKeyFun == "")
            {
                for (var i = 0; i < count; i++)
                {
                    var replaceCell = wkSheet.Cells[writeRow + i + 1, cellCol];
                    //判断使用通用替换还是专属替换
                    string replaceValueCheck;
                    if (!string.IsNullOrEmpty(specialValue[0, 0]))
                    {
                        replaceValueCheck = $"{_replaceValues}#{specialValue[0, 0]}";
                    }
                    else
                    {
                        replaceValueCheck = _replaceValues;
                    }
                    string[] baseParts = replaceValueCheck.Split("#");
                    var cellValue = replaceCell.Value?.ToString() ?? "";

                    int partCount = 0;
                    foreach (var item in baseParts)
                    {
                        var parts = item.Split("-");
                        var replaceValue = parts[0];
                        //备注的全量替换
                        if (replaceValue.Contains("***"))
                        {
                            replaceValue = baseParts[partCount].Replace("***", "");
                            cellValue = replaceValue;
                        }
                        else
                        {
                            var pattern = parts[1];
                            if (cellValue != null)
                                cellValue = Regex.Replace(cellValue, pattern, replaceValue);
                        }
                        partCount++;
                    }
                    replaceCell.Value = cellValue;
                }
            }
            //自定义修改（修改方法）
            else if (cellKey != "")
            {
                var excelFileFixKey = PubMetToExcel.FindSourceCol(wkSheet, 2, cellKey);
                for (var i = 0; i < count; i++)
                {
                    var cellSource = wkSheet.Cells[startRowSource + i, excelFileFixKey];
                    var rowId = wkSheet.Cells[startRowSource + i, 2];
                    var cellFix = wkSheet.Cells[writeRow + 1 + i, excelFileFixKey];

                    string cellFixValue;
                    //固定值
                    string baseValue = excelKeyFun ?? "";
                    if (baseValue.Contains("***"))
                    {
                        baseValue = baseValue.Replace("***", "");
                        cellFixValue = baseValue;
                    }
                    //自增值
                    else
                    {
                        if (cellSource.Value == null)
                            continue;

                        if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0")
                            continue;

                        var fixValueList = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyFun);
                        cellFixValue = ExcelDataAutoInsert.StringRegPlace(
                            cellSource.Value.ToString(),
                            fixValueList,
                            _addValue
                        );
                    }

                    if (cellFixValue == "^error^")
                    {
                        string errorExcelLog =
                            excelName + "#" + rowId.Value + "#【修改模式】#[" + cellKey + "]字段方法写错";
                        errorList.Add((cellKey, errorExcelLog, excelName));
                    }

                    cellFix.Value = double.TryParse(cellFixValue, out double number)
                        ? number
                        : cellFixValue;
                }
            }
            //备注修改（替换）
            if (cellKey.Contains("#") || cellKey == "comment")
            {
                for (var i = 0; i < count; i++)
                {
                    var replaceCell = wkSheet.Cells[writeRow + i + 1, cellCol];

                    foreach (var comment in _commentValue)
                    {
                        var replaceCellValue = replaceCell.Value?.ToString() ?? "";
                        if (replaceCellValue.Contains(comment.Key))
                        {
                            var replaceComment = comment.Value[0].Item1[0, 0].ToString();
                            //空值不替换
                            if (replaceComment != "")
                            {
                                replaceCellValue = replaceCellValue.Replace(
                                    comment.Key,
                                    replaceComment
                                );
                                replaceCell.Value = replaceCellValue;
                            }
                        }
                        //else
                        //{
                        //    replaceCell.Value = "*?*" + replaceCellValue ;
                        //}
                    }
                }
            }
        }

        return errorList;
    }

    private static (List<string>, int) GetElementIdGroup(
        dynamic excelName,
        ExcelWorksheet wkSheet,
        dynamic modelIdnew,
        bool isDelete = false
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var lastRow = 0;
        for (var excelMulti = 0; excelMulti < modelIdnew[excelName].Count; excelMulti++)
        {
            var startValue = modelIdnew[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelIdnew[excelName][excelMulti].Item1[1, 0].ToString();
            var startRowSource = PubMetToExcel.FindSourceRow(wkSheet, 2, startValue);
            var endRowSource = PubMetToExcel.FindSourceRow(wkSheet, 2, endValue);
            if (startRowSource == -1 || endRowSource == -1)
            {
                if (isDelete)
                {
                    var writeIdList2 = new List<string> { startValue + "#" + endValue };
                    return (writeIdList2, -1);
                }
            }
            if (endRowSource < startRowSource)
            {
                return ([$"{endValue}-有重复值"], -9527);
            }
            var count = endRowSource - startRowSource + 1;
            if (isDelete)
            {
                wkSheet.DeleteRow(startRowSource, count);
            }

            if (lastRow < endRowSource)
                lastRow = endRowSource;
        }
        return (["查询完毕：正确"], lastRow);
    }
}
//纯Epplus，不使用Interop，无法针对激活Sheet使用
public static class ExcelDataAutoInsertCopyMulti
{
    public static void SearchData(dynamic isMulti)
    {
        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        var colorCell = sheet.Cells[6, 1];
        PubMetToExcel.GetCellBackgroundColor(colorCell);
        ErrorLogCtp.DisposeCtp();
        var rowCount = 2;
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(
            data,
            sheetNameCol,
            modelIdNewCol,
            rowCount
        );
        var errorExcelList = new List<List<(string, string, string)>>();
        var excelCount = 1;
        var diffList = new List<(string, string, string)>();
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        foreach (var key in modelIdNew)
        {
            var excelName = key.Key;
            var targetExcelPath =
                excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
            List<(string, string, string)> errorList = PubMetToExcel.SetExcelObjectEpPlus(
                targetExcelPath,
                excelName,
                out ExcelWorksheet targetSheet,
                out ExcelPackage _
            );
            if (errorList.Count != 0) { }

            errorList = PubMetToExcel.SetExcelObjectEpPlus(
                excelPath,
                excelName,
                out ExcelWorksheet sourceSheet,
                out ExcelPackage _
            );
            if (errorList.Count != 0) { }

            for (var excelMulti = 0; excelMulti < modelIdNew[excelName].Count; excelMulti++)
            {
                var startValue = modelIdNew[excelName][excelMulti].Item1[0, 0].ToString();
                var endValue = modelIdNew[excelName][excelMulti].Item1[1, 0].ToString();
                var startRowSource = PubMetToExcel.FindSourceRow(sourceSheet, 2, startValue);
                var endRowSource = PubMetToExcel.FindSourceRow(sourceSheet, 2, endValue);
                var startRowTarget = PubMetToExcel.FindSourceRow(targetSheet, 2, startValue);
                var endRowTarget = PubMetToExcel.FindSourceRow(targetSheet, 2, endValue);
                if (endRowSource - startRowSource > endRowTarget - startRowTarget)
                    for (int i = startRowSource; i <= endRowSource; i++)
                    {
                        var cellSourceValue = sourceSheet.Cells[i, 2].Value.ToString();
                        var resultValue = "";
                        var resultRow = "";
                        for (int j = startRowTarget; j <= endRowTarget; j++)
                        {
                            var cellTargetValue = targetSheet.Cells[j, 2].Value.ToString();
                            if (cellSourceValue != cellTargetValue)
                            {
                                resultValue = cellTargetValue;
#pragma warning disable CA1305
                                resultRow = j.ToString();
#pragma warning restore CA1305
                            }
                        }

                        if (resultValue != "")
                            diffList.Add((excelPath + @"\" + excelName, resultRow, resultValue));
                    }
                else
                    for (int i = startRowTarget; i <= endRowTarget; i++)
                    {
                        var cellTargetValue = targetSheet.Cells[i, 2].Value.ToString();
                        var resultValue = "";
                        var resultRow = "";
                        for (int j = startRowSource; j <= endRowSource; j++)
                        {
                            var cellSourceValue = sourceSheet.Cells[j, 2].Value.ToString();
                            if (cellSourceValue != cellTargetValue)
                            {
                                resultValue = cellSourceValue;
#pragma warning disable CA1305
                                resultRow = j.ToString();
#pragma warning restore CA1305
                            }
                        }

                        if (resultValue != "")
                            diffList.Add(
                                (targetExcelPath + @"\" + excelName, resultRow, resultValue)
                            );
                    }
            }

            NumDesAddIn.App.StatusBar =
                "遍历表格" + "<" + excelCount + "/" + modelIdNew.Count + ">" + excelName;
            errorExcelList.Add(errorList);
            excelCount++;
        }

        diffList = diffList.Distinct().ToList().ToList();
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            dynamic tempWorkbook;
            try
            {
                tempWorkbook = NumDesAddIn.App.Workbooks.Open(excelPath + @"\#合并表格数据缓存.xlsx");
            }
            catch
            {
                tempWorkbook = NumDesAddIn.App.Workbooks.Add();
                tempWorkbook.SaveAs(excelPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
            }

            var tempSheet = tempWorkbook.Sheets["Sheet1"];
            Range usedRange = tempSheet.UsedRange;
            usedRange.ClearContents();
            var tempDataArray = new string[diffList.Count, 4];
            for (var i = 0; i < diffList.Count; i++)
            {
                tempDataArray[i, 0] = diffList[i].Item1;
                tempDataArray[i, 1] = "Sheet1";
                tempDataArray[i, 2] = "B" + diffList[i].Item2;
                tempDataArray[i, 3] = diffList[i].Item3;
            }

            var tempDataRange = tempSheet.Range[
                tempSheet.Cells[2, 2],
                tempSheet.Cells[
                    2 + tempDataArray.GetLength(0) - 1,
                    2 + tempDataArray.GetLength(1) - 1
                ]
            ];
            tempDataRange.Value = tempDataArray;
            tempWorkbook.Save();
            NumDesAddIn.App.Visible = true;
            NumDesAddIn.App.StatusBar = "完成统计";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static void MergeData(dynamic isMulti)
    {
        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        ErrorLogCtp.DisposeCtp();
        var rowCount = 2;
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(
            data,
            sheetNameCol,
            modelIdNewCol,
            rowCount
        );
        var errorExcelList = new List<List<(string, string, string)>>();
        var excelCount = 1;
        foreach (var key in modelIdNew)
        {
            var excelName = key.Key;
            List<(string, string, string)> error = AutoCopyData(
                modelIdNew,
                excelName,
                excelPath,
                cellColor
            );
            NumDesAddIn.App.StatusBar =
                "写入数据" + "<" + excelCount + "/" + modelIdNew.Count + ">" + excelName;
            errorExcelList.Add(error);
            excelCount++;
        }

        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            NumDesAddIn.App.StatusBar = "完成写入";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    private static List<(string, string, string)> AutoCopyData(
        dynamic modelIdNew,
        dynamic excelName,
        dynamic excelPath,
        dynamic cellColor
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorList = new List<(string, string, string)>();
        var targetExcelPath = "";
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        if (mergePathList.Count <= 1)
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            Process.Start(filePath);
            return errorList;
        }

        if (
            mergePathList[0] == ""
            || mergePathList[1] == ""
            || mergePathList[1] == mergePathList[0]
        )
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            Process.Start(filePath);
        }
        else
        {
            targetExcelPath = excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        }

        if (targetExcelPath == "")
            return errorList;

        errorList = PubMetToExcel.SetExcelObjectEpPlus(
            targetExcelPath,
            excelName,
            out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel
        );
        if (errorList.Count != 0)
            return errorList;
        errorList = PubMetToExcel.SetExcelObjectEpPlus(
            excelPath,
            excelName,
            out ExcelWorksheet sourceSheet,
            out ExcelPackage _
        );
        if (errorList.Count != 0)
            return errorList;
        foreach (var cell in targetSheet.Cells)
        {
            if (cell.Formula is not { Length: > 0 })
                continue;
            errorList.Add((excelName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
            return errorList;
        }

        for (var excelMulti = 0; excelMulti < modelIdNew[excelName].Count; excelMulti++)
        {
            var startValue = modelIdNew[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelIdNew[excelName][excelMulti].Item1[1, 0].ToString();
            var startRowSource = PubMetToExcel.FindSourceRow(sourceSheet, 2, startValue);
            string errorExcelLog;
            if (startRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
                errorList.Add((startValue, errorExcelLog, excelName));
                return errorList;
            }

            var endRowSource = PubMetToExcel.FindSourceRow(sourceSheet, 2, endValue);
            if (endRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]未找到(序号出错)";
                errorList.Add((endValue, errorExcelLog, excelName));
                return errorList;
            }

            if (endRowSource - startRowSource < 0)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]起始、终结ID顺序反了";
                errorList.Add((endValue, errorExcelLog, excelName));
                return errorList;
            }

            var targetMaxCol = targetSheet.Dimension.Columns;
            var sourceMaxCol = sourceSheet.Dimension.Columns;
            var targetRangeTitle = (object[,])targetSheet.Cells[2, 1, 2, targetMaxCol].Value;
            var sourceRangeTitle = (object[,])sourceSheet.Cells[2, 1, 2, sourceMaxCol].Value;
            var sourceRangeValue = (object[,])
                sourceSheet.Cells[startRowSource, 1, endRowSource, sourceMaxCol].Value;
            var targetRowList = PubMetToExcel.MergeExcel(
                sourceRangeValue,
                targetSheet,
                targetRangeTitle,
                sourceRangeTitle
            );
            for (var i = 0; i < targetRowList.Count; i++)
            {
                var cellTarget = targetSheet.Cells[
                    targetRowList[i],
                    1,
                    targetRowList[i],
                    targetMaxCol
                ];
                var isColorCell = targetSheet.Cells[targetRowList[i], 2];
                if (isColorCell.Style.Fill.BackgroundColor.Rgb == null)
                {
                    cellTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cellTarget.Style.Fill.BackgroundColor.SetColor(cellColor);
                }
            }
        }

        targetExcel.Save();
        targetSheet.Dispose();
        return errorList;
    }

    public static void RightClickMergeData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        NumDesAddIn.App.StatusBar = false;
        var sw = new Stopwatch();
        sw.Start();

        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var excelName = indexWk.Name;
        ErrorLogCtp.DisposeCtp();
        var errorExcelList = new List<List<(string, string, string)>>();
        List<(string, string, string)> error = AutoCopyDataRight(
            NumDesAddIn.App,
            excelPath,
            excelName,
            sheet
        );
        errorExcelList.Add(error);
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            sw.Stop();
            var ts1 = sw.ElapsedMilliseconds;
            NumDesAddIn.App.StatusBar = "完成写入：" + ts1;
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);

        sw.Stop();
        var ts2 = sw.ElapsedMilliseconds;
        NumDesAddIn.App.StatusBar = "完成写入：" + ts2;
        Marshal.ReleaseComObject(sheet);
        Marshal.ReleaseComObject(indexWk);
    }

    private static List<(string, string, string)> AutoCopyDataRight(
        dynamic app,
        dynamic excelPath,
        dynamic excelName,
        dynamic sheet
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorList = new List<(string, string, string)>();
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);

        if (
            mergePathList.Count <= 1
            || mergePathList[0] == ""
            || mergePathList[1] == ""
            || mergePathList[1] == mergePathList[0]
            || !mergePathList[0].Contains("Tables")
            || !mergePathList[1].Contains("Tables")
        )
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            Process.Start(filePath);
            return errorList;
        }

        var targetExcelPath = excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        if (targetExcelPath == "")
            return errorList;

        errorList = PubMetToExcel.SetExcelObjectEpPlus(
            targetExcelPath,
            excelName,
            out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel
        );
        if (errorList.Count != 0)
            return errorList;
        errorList = PubMetToExcel.SetExcelObjectEpPlus(
            excelPath,
            excelName,
            out ExcelWorksheet sourceSheet,
            out ExcelPackage _
        );
        if (errorList.Count != 0)
            return errorList;
        foreach (var cell in targetSheet.Cells)
            if (cell.Formula is { Length: > 0 })
            {
                errorList.Add((excelName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
                return errorList;
            }

        var targetMaxCol = targetSheet.Dimension.Columns;
        var sourceMaxCol = sourceSheet.Dimension.Columns;
        var sourceRangeTitle = sourceSheet.Cells[2, 1, 2, sourceMaxCol];
        var targetRangeTitle = targetSheet.Cells[2, 1, 2, targetMaxCol];
        var selectRange = app.Selection;

        if (selectRange.Cells.Count > 0)
        {
            int minRow = selectRange.Row;
            int maxRow = selectRange.Row + selectRange.Rows.Count - 1;
            var sourceRangeValue = (object[,])
                sourceSheet.Cells[minRow, 1, maxRow, sourceMaxCol].Value;
            var sourceRangeValueTitle = (object[,])sourceRangeTitle.Value;
            var targetRangeValueTitle = (object[,])targetRangeTitle.Value;
            var targetRowList = PubMetToExcel.MergeExcel(
                sourceRangeValue,
                targetSheet,
                targetRangeValueTitle,
                sourceRangeValueTitle
            );
            var colorCell = sheet.Cells[minRow, 2];
            var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
            for (var i = 0; i < targetRowList.Count; i++)
            {
                var cellTarget = targetSheet.Cells[
                    targetRowList[i],
                    1,
                    targetRowList[i],
                    targetMaxCol
                ];
                var isColorCell = targetSheet.Cells[targetRowList[i], 2];
                if (isColorCell.Style.Fill.BackgroundColor.Rgb == null)
                {
                    cellTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cellTarget.Style.Fill.BackgroundColor.SetColor(cellColor);
                }
            }
        }

        targetExcel.Save();
        targetSheet.Dispose();


        return errorList;
    }

    public static void RightClickMergeDataCol(CommandBarButton ctrl, ref bool cancelDefault)
    {
        NumDesAddIn.App.StatusBar = false;
        var sw = new Stopwatch();
        sw.Start();

        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var excelName = indexWk.Name;
        ErrorLogCtp.DisposeCtp();
        var errorExcelList = new List<List<(string, string, string)>>();
        var error = AutoCopyDataRightCol(NumDesAddIn.App, excelPath, excelName);
        errorExcelList.Add(error);
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            sw.Stop();
            var ts1 = sw.ElapsedMilliseconds;
            NumDesAddIn.App.StatusBar = "完成写入：" + ts1;
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);

        sw.Stop();
        var ts2 = sw.ElapsedMilliseconds;
        NumDesAddIn.App.StatusBar = "完成写入：" + ts2;
    }

    private static List<(string, string, string)> AutoCopyDataRightCol(
        dynamic app,
        dynamic excelPath,
        dynamic excelName
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorList = new List<(string, string, string)>();
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);

        if (
            mergePathList.Count <= 1
            || mergePathList[0] == ""
            || mergePathList[1] == ""
            || mergePathList[1] == mergePathList[0]
            || !mergePathList[0].Contains("Tables")
            || !mergePathList[1].Contains("Tables")
        )
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            Process.Start(filePath);
            return errorList;
        }

        var targetExcelPath = excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        if (targetExcelPath == "")
            return errorList;

        errorList = PubMetToExcel.SetExcelObjectEpPlus(
            targetExcelPath,
            excelName,
            out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel
        );
        if (errorList.Count != 0)
            return errorList;
        errorList = PubMetToExcel.SetExcelObjectEpPlus(
            excelPath,
            excelName,
            out ExcelWorksheet sourceSheet,
            out ExcelPackage _
        );
        if (errorList.Count != 0)
            return errorList;
        foreach (var cell in targetSheet.Cells)
            if (cell.Formula is { Length: > 0 })
            {
                errorList.Add((excelName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
                return errorList;
            }

        var targetMaxRow = targetSheet.Dimension.Rows;
        var sourceMaxRow = sourceSheet.Dimension.Rows;
        var sourceRangeTitle = sourceSheet.Cells[2, 2, targetMaxRow, 2];
        var targetRangeTitle = targetSheet.Cells[2, 2, sourceMaxRow, 2];
        var selectRange = app.Selection;

        if (selectRange.Cells.Count > 0)
        {
            int minCol = selectRange.Column;
            int maxCol = selectRange.Column + selectRange.Column.Count - 1;
            var sourceRangeValue = (object[,])
                sourceSheet.Cells[1, minCol, sourceMaxRow, maxCol].Value;
            var sourceRangeValueTitle = (object[,])sourceRangeTitle.Value;
            var targetRangeValueTitle = (object[,])targetRangeTitle.Value;
            PubMetToExcel.MergeExcelCol(
                sourceRangeValue,
                targetSheet,
                targetRangeValueTitle,
                sourceRangeValueTitle
            );
        }

        targetExcel.Save();
        targetSheet.Dispose();
        return errorList;
    }
}

public static class ExcelDataAutoInsertActivityServer
{
    private const double SecondsInADay = 86400;
    private const double OneMinuteInDays = 60 / SecondsInADay;

    public static void Source(bool isNames)
    {
        var indexWk = NumDesAddIn.App.ActiveWorkbook;

        var sourceSheet = indexWk.Worksheets["运营排期"];
        var targetSheet = indexWk.Worksheets["Sheet1"];
        var fixSheet = indexWk.Worksheets["活动模板"];
        var lifeTypeSheet = indexWk.Worksheets["生命周期"];

        var fixData = PubMetToExcel.ExcelDataToList(fixSheet);
        var fixTitle = fixData.Item1;
        List<List<object>> fixDataList = fixData.Item2;
        //删除活动名或者活动id列为空的数据
        fixDataList = fixDataList.Where(row => row[0] != null && row[1] != null).ToList();

        var fixNames = fixTitle.IndexOf("活动名称");
        var fixIds = fixTitle.IndexOf("活动id");
        var fixPush = fixTitle.IndexOf("前端可获取活动时间");
        //var fixPushEnds = fixTitle.IndexOf("停止向前端发送活动时间");
        //var fixPreHeats = fixTitle.IndexOf("预热期开始时间");
        //var fixOpens = fixTitle.IndexOf("活动开启时间");
        //var fixEnds = fixTitle.IndexOf("活动结束时间");
        var fixCloses = fixTitle.IndexOf("活动关闭时间");
        var isActGroup = fixTitle.IndexOf("是否活动组");
        var openCondition = fixTitle.IndexOf("活动开启条件");
        var lifeType = fixTitle.IndexOf("生命周期类型");

        var lifeTypeData = PubMetToExcel.ExcelDataToList(lifeTypeSheet);
        var lifeTypeTitle = lifeTypeData.Item1;
        List<List<object>> lifeTypeDataList = lifeTypeData.Item2;
        var lifeTypeIndex = lifeTypeTitle.IndexOf("类型");
        var lifeTypeValue = lifeTypeTitle.IndexOf("内容");

        var sourceMaxCol = sourceSheet.UsedRange.Columns.Count;
        var sourceMaxRow = sourceSheet.UsedRange.Rows.Count;
        var sourceRange = sourceSheet.Range[
            sourceSheet.Cells[3, 5],
            sourceSheet.Cells[sourceMaxRow, sourceMaxCol]
        ];
        var sourceDateRange = sourceSheet.Range[
            sourceSheet.Cells[3, 3],
            sourceSheet.Cells[sourceMaxRow, 3]
        ];
        var sourceOutRange = sourceSheet.Range[
            sourceSheet.Cells[2, 5],
            sourceSheet.Cells[2, sourceMaxCol]
        ];

        int nameOrId = isNames ? fixNames : fixIds;
        string nameOrIdString = isNames ? "活动名" : "活动ID";

        Array sourceDataArr = sourceDateRange.Value2;
        var sourceData = new List<(string, double, double, int, int, int, string)>();

        for (int col = 1; col <= sourceMaxCol - 3 + 1; col++)
        {
            for (int row = 1; row <= sourceMaxRow - 3 + 1; row++)
            {
                var cell = sourceRange[row, col];

                // 过滤已删除活动
                bool hasStrikethrough = cell.Font.Strikethrough;
                if (hasStrikethrough)
                    continue;

                var cellOutValue = sourceOutRange[1, col].Value2?.ToString() ?? "";
                if (cellOutValue != "#导出")
                    continue;

                if (cell.MergeCells)
                {
                    var mergeRange = cell.MergeArea;
                    if (cell.Address == mergeRange.Cells[1, 1].Address)
                    {
                        var mergeValue = mergeRange.Cells[1, 1].Value2;
                        if (mergeValue == null)
                            continue;
                        var activityName = mergeValue.ToString();
                        var activityCondition = "";
                        if (activityName.Contains("："))
                        {
                            activityName = mergeValue.ToString().Split("：")[1];
                            activityCondition = mergeValue.ToString().Split("：")[0];
                        }

                        sourceData.Add(
                            (
                                activityName,
                                (double)sourceDataArr.GetValue(mergeRange.Row - 2, 1),
                                (double)
                                    sourceDataArr.GetValue(
                                        mergeRange.Row + mergeRange.Rows.Count - 3,
                                        1
                                    ),
                                mergeRange.Column,
                                mergeRange.Row,
                                mergeRange.Row + mergeRange.Rows.Count - 1,
                                activityCondition
                            )
                        );
                    }
                }
                else if (cell.Value != null)
                {
                    var activityName = cell.Value.ToString();
                    var activityCondition = "";
                    if (activityName.Contains("："))
                    {
                        activityName = cell.Value.ToString().Split("：")[1];
                        activityCondition = cell.Value.ToString().Split("：")[0];
                    }
                    sourceData.Add(
                        (
                            activityName,
                            (double)sourceDataArr.GetValue(cell.Row - 2, 1),
                            (double)sourceDataArr.GetValue(cell.Row + cell.Rows.Count - 3, 1),
                            cell.Column,
                            cell.Row,
                            cell.Row + cell.Rows.Count - 1,
                            activityCondition
                        )
                    );
                }
            }
        }

        var targetDataList = new List<List<string>>();
        var errorLog = "";
        var unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        foreach (var a in sourceData)
        {
            var fixDataMatch = fixDataList.FirstOrDefault(b => b[nameOrId].ToString() == a.Item1);
            if (fixDataMatch == null)
            {
                var activeName = a.Item1;
                if (a.Item7 != "")
                {
                    activeName = $"{a.Item7}：{a.Item1}";
                }
                errorLog += $"运营排期-未找到-活动模版【{nameOrIdString}】：{activeName}\r\n";
                targetDataList.Add(
                    [
                        "targetId",
                        a.Item1,
                        "targetPushTimeString",
                        "targetPushTimeLong",
                        "targetPushEndTimeString",
                        "targetPushEndTimeLong",
                        "targetPreHeatTimeString",
                        "targetPreHeatTimeLong",
                        "targetOpenTimeString",
                        "targetOpenTimeLong",
                        "targetEndTimeString",
                        "targetEndTimeLong",
                        "targetCloseTimeString",
                        "targetCloseTimeLong",
                        "targetActGroup",
                        "targetOpenCondition",
                        "targetLifeType"
                    ]
                );
                continue;
            }

            var sourceStartTimeLong = (long)
                (DateTime.FromOADate(a.Item2).ToUniversalTime() - unixEpoch).TotalSeconds;
            var sourceEndTimeLong = (long)
                (
                    DateTime
                        .FromOADate(a.Item2 + a.Item6 - a.Item5 + 1 - OneMinuteInDays)
                        .ToUniversalTime() - unixEpoch
                ).TotalSeconds;

            string ConvertToDateString(double oaDate, double hoursOffset)
            {
                return DateTime
                    .FromOADate(oaDate)
                    .AddHours(hoursOffset * 24)
                    .ToString(CultureInfo.InvariantCulture);
            }

            long ConvertToUnixTime(long baseTime, double hoursOffset)
            {
                return baseTime + (long)(hoursOffset * 24 * 3600);
            }

            var targetId = fixDataMatch[fixIds].ToString();
            var targetName = a.Item1;
            if (a.Item7 != "")
            {
                targetName = $"{a.Item7}：{a.Item1}";
            }
            var targetPushTimeString = ConvertToDateString(a.Item2, fixDataMatch[fixPush]);
            var targetPushTimeLong = ConvertToUnixTime(sourceStartTimeLong, fixDataMatch[fixPush]);
            var targetPushEndTimeString = ConvertToDateString(
                a.Item2 + a.Item6 - a.Item5 + 1 - OneMinuteInDays,
                0
            //fixDataMatch[fixPushEnds]
            );
            var targetPushEndTimeLong = ConvertToUnixTime(
                sourceEndTimeLong,
                0
            //fixDataMatch[fixPushEnds]
            );
            var targetPreHeatTimeString = ConvertToDateString(
                a.Item2,
                0 /*fixDataMatch[fixPreHeats]*/
            );
            var targetPreHeatTimeLong = ConvertToUnixTime(
                sourceStartTimeLong,
                0
            //fixDataMatch[fixPreHeats]
            );
            var targetOpenTimeString = ConvertToDateString(
                a.Item2,
                0 /*fixDataMatch[fixOpens]*/
            );
            var targetOpenTimeLong = ConvertToUnixTime(
                sourceStartTimeLong,
                0 /*fixDataMatch[fixOpens]*/
            );
            var targetEndTimeString = ConvertToDateString(
                a.Item3 - OneMinuteInDays,
                0
                    /*fixDataMatch[fixEnds] */+ 1
            );
            var targetEndTimeLong = ConvertToUnixTime(
                sourceEndTimeLong,
                0 /*fixDataMatch[fixEnds]*/
            );
            var targetCloseTimeString = ConvertToDateString(
                a.Item3 - OneMinuteInDays,
                fixDataMatch[fixCloses] + 1
            );
            var targetCloseTimeLong = ConvertToUnixTime(sourceEndTimeLong, fixDataMatch[fixCloses]);
            var targetActGroup = fixDataMatch[isActGroup].ToString();
            var targetOpenCondition = fixDataMatch[openCondition]?.ToString() ?? "";
            if (targetOpenCondition == "\"{}\"" || targetOpenCondition == "")
            {
                if (a.Item7 != "")
                {
                    targetOpenCondition = "\"{{26,{" + a.Item7.Replace("、", ",") + "}}}\"";
                }
            }
            var targetLifeType = fixDataMatch[lifeType];
            string targetLifeValue;
            if (targetLifeType == null)
            {
                targetLifeValue = "";
            }
            else
            {
                var lifeTypeMatch = lifeTypeDataList.FirstOrDefault(l =>
                    l[lifeTypeIndex].ToString() == targetLifeType.ToString()
                );
                if (lifeTypeMatch == null)
                {
                    var activeName = a.Item1;
                    if (a.Item7 != "")
                    {
                        activeName = $"{a.Item7}：{a.Item1}";
                    }
                    errorLog +=
                        $"运营排期-活动模版【{nameOrIdString}】："
                        + activeName
                        + $"**生命周期类型错误[{targetLifeType}]，搜索不到\r\n";
                    targetLifeValue = "targetLifeValue";
                }
                else
                {
                    targetLifeValue = lifeTypeMatch[lifeTypeValue]?.ToString() ?? "";
                }
            }
            targetDataList.Add(
                [
                    targetId,
                    targetName,
                    targetPushTimeString,
                    targetPushTimeLong.ToString(),
                    targetPushEndTimeString,
                    targetPushEndTimeLong.ToString(CultureInfo.InvariantCulture),
                    targetPreHeatTimeString,
                    targetPreHeatTimeLong.ToString(),
                    targetOpenTimeString,
                    targetOpenTimeLong.ToString(),
                    targetEndTimeString,
                    targetEndTimeLong.ToString(CultureInfo.InvariantCulture),
                    targetCloseTimeString,
                    targetCloseTimeLong.ToString(CultureInfo.InvariantCulture),
                    targetActGroup,
                    targetOpenCondition,
                    targetLifeValue
                ]
            );
        }

        if (!string.IsNullOrEmpty(errorLog))
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtp(errorLog);
            MessageBox.Show(@"有活动找不到，查看错误日志");
            sourceSheet.Select();
        }
        else
        {
            targetSheet.Select();
        }
        var targetStartCol = 2;
        var targetStartRow = 5;
        var targetRangeOld = targetSheet.Range[
            targetSheet.Cells[targetStartRow, targetStartCol],
            targetSheet.Cells[targetSheet.UsedRange.Rows.Count, targetSheet.UsedRange.Columns.Count]
        ];
        targetRangeOld.Value = null;

        var rows = targetDataList.Count;
        var columns = targetDataList[0].Count;
        var targetDataArr = new string[rows, columns];
        for (var i = 0; i < rows; i++)
        {
            for (var j = 0; j < columns; j++)
            {
                targetDataArr[i, j] = targetDataList[i][j];
            }
        }

        var targetRange = targetSheet.Range[
            targetSheet.Cells[targetStartRow, targetStartCol],
            targetSheet.Cells[
                targetStartRow + targetDataArr.GetLength(0) - 1,
                targetStartCol + targetDataArr.GetLength(1) - 1
            ]
        ];
        targetRange.Value = targetDataArr;
    }
}

public class ExcelDataAutoInsertNumChanges
{
    private string _excelPath;

    private Dictionary<string, (List<object>, List<List<object>>)> GetNumChangesData(
        int startRow,
        dynamic indexSheet,
        dynamic startValue,
        dynamic workBook
    )
    {
        var usedRange = indexSheet.UsedRange;
        var rowMax = usedRange.Rows.Count;
        var colMax = usedRange.Columns.Count;

        var dataList = new Dictionary<string, (List<object>, List<List<object>>)>();

        for (int col = startValue.Item2 + 2; col <= colMax; col++)
        {
            var isCurrentCol = (col - startValue.Item2) % 4;
            if (isCurrentCol == 2)
            {
                var startCell = indexSheet.Cells[startRow + 2, col];
                var endCell = indexSheet.Cells[rowMax, col + 3];
                var dataRange = indexSheet.Range[startCell, endCell];

                var startHeadCell = indexSheet.Cells[startRow + 1, col];
                var endHeadCell = indexSheet.Cells[startRow + 1, col + 3];
                var headRange = indexSheet.Range[startHeadCell, endHeadCell];

                var data = workBook.Read(dataRange, headRange, 1);

                var targetBookRange = indexSheet.Cells[startRow, col];
                var targetBookRangeName = targetBookRange.Value?.ToString();

                if (targetBookRangeName != null)
                    dataList.Add(targetBookRangeName, data);
            }
        }

        return dataList;
    }

    public void SetNumChangesData(Dictionary<string, (List<object>, List<List<object>>)> data)
    {
        if (_excelPath == null)
        {
            var wk = NumDesAddIn.App.ActiveWorkbook;
            _excelPath = wk.Path;
        }
        foreach (var eachExcelData in data)
        {
            var workBookName = eachExcelData.Key;
            var excelObj = new ExcelDataByEpplus();
            excelObj.GetExcelObj(_excelPath, workBookName);
            if (excelObj.ErrorList.Count > 0)
            {
                MessageBox.Show($"{workBookName}不存在，少个#?");
                return;
            }

            var changeValueCount = 0;

            var sheetTarget = excelObj.Sheet;
            var excelTarget = excelObj.Excel;
            var keyIndex = eachExcelData.Value.Item1[0].ToString();
            var keyIndexRowCount = eachExcelData.Value.Item2.Count;
            var keyIndexCol = excelObj.FindFromCol(sheetTarget, 2, keyIndex);

            for (int j = 1; j < eachExcelData.Value.Item1.Count; j++)
            {
                var keyTarget = eachExcelData.Value.Item1[j].ToString();
                if (keyTarget != null && keyTarget.Contains("$"))
                {
                    continue;
                }
                var keyTargetCol = excelObj.FindFromCol(sheetTarget, 2, keyTarget);
                if (keyIndexCol == -1 || keyTargetCol == -1)
                {
                    MessageBox.Show(workBookName + "*找不到字段");
                    return;
                }

                for (var i = 0; i < keyIndexRowCount; i++)
                {
                    var keyIndexValue = eachExcelData.Value.Item2[i][0]?.ToString();
                    var keyTargetValue = eachExcelData.Value.Item2[i][j]?.ToString();
                    if (keyIndexValue != null && keyTargetValue != null && keyIndexValue != "")
                    {
                        var keyIndexRow = excelObj.FindFromRow(
                            sheetTarget,
                            keyIndexCol,
                            keyIndexValue
                        );
                        LogDisplay.RecordLine(
                            "[{0}] , {1}",
                            DateTime.Now.ToString(CultureInfo.InvariantCulture),
                            keyIndexValue
                        );
                        if (keyIndexRow == -1)
                        {
                            MessageBox.Show($"{workBookName} 找不到Id：{keyIndexValue}");
                            LogDisplay.RecordLine(
                                "[{0}] , {1}",
                                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                                $"{workBookName} 找不到Id：{keyIndexValue}"
                            );
                            return;
                        }
                        var baseValue = sheetTarget
                            .Cells[keyIndexRow, keyTargetCol]
                            .Value?.ToString();
                        // ReSharper disable once StringLiteralTypo
                        if (baseValue != keyTargetValue && keyTargetValue != "nofix")
                        {
                            sheetTarget.Cells[keyIndexRow, keyTargetCol].Value = double.TryParse(
                                keyTargetValue,
                                out double number
                            )
                                ? number
                                : keyTargetValue;
                            changeValueCount++;
                        }
                    }
                }
            }

            if (changeValueCount > 0)
                excelTarget.Save();
        }
    }

    public void OutDataIsAll()
    {
        var workBook = new ExcelDataByVsto();
        workBook.GetExcelObj();
        var indexSheet = workBook.ActiveSheet;
        var indexRange = indexSheet.UsedRange;
        var startValue = workBook.FindValue(indexRange, "*自动填表*");
        int startRow = startValue.Item1;
        if (startRow == -1)
        {
            MessageBox.Show("表格中找不到【*自动填表*】");
        }
        var activityRankRange = indexSheet.Cells[startRow - 1, startValue.Item2 + 1];
        var activityRankCountRange = indexSheet.Cells[startRow - 2, startValue.Item2 + 1];
        var activityRank = (int)activityRankRange.Value;
        var activityRankCount = (int)activityRankCountRange.Value;
        _excelPath = workBook.ActiveWorkbookPath;

        var tips = MessageBox.Show(
            "是否导出全部活动数据（Y：全部；N：当前）",
            "确认",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        );
        if (tips == MessageBoxResult.Yes)
        {
            for (var i = activityRank; i <= activityRankCount; i++)
            {
                activityRankRange.Value = i;
                var data = GetNumChangesData(startRow, indexSheet, startValue, workBook);
                SetNumChangesData(data);
            }

            activityRankRange.Value = activityRank;
        }
        else
        {
            var data = GetNumChangesData(startRow, indexSheet, startValue, workBook);
            SetNumChangesData(data);
        }
    }
}

public static class AutoInsertExcelDataModelCreat
{
    public static void InsertModelData(dynamic wk)
    {
        var sheet = wk.ActiveSheet;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        List<List<object>> data = sheetData.Item2;

        //写入数据
        //通配符信息
        var exportWildcardData = new Dictionary<string, string>();
        if (exportWildcardData == null)
            throw new ArgumentNullException(nameof(exportWildcardData));
        foreach (var row in data)
        {
            if (row.Count >= 2) // 确保有至少两列
            {
                var key = row[0]?.ToString();
                var value = row[1]?.ToString();
                if (key == null || value == null)
                    continue;
                exportWildcardData.TryAdd(key, value);
            }
        }
        //获取模版ListObject数据，并替换数据

        var modelSheet = sheet;
        if (!sheet.Name.Contains("LTE皮肤"))
        {
            MessageBox.Show("当前表格不是【LTE皮肤……】类表格，不能使用该功能");
            return;
        }
        var modelListObjects = modelSheet.ListObjects;
        var modelValueAll = new Dictionary<string, (List<object>, List<List<object>>)>();

        foreach (ListObject list in modelListObjects)
        {
            var modelName = list.Name;
            if (modelName.Contains("Dollar"))
            {
                modelName = modelName.Replace("Dollar", "$");
                modelName = modelName.Replace("_", "##");
            }
            //截取.xlsx之前的字符
            modelName = modelName.Substring(0, modelName.IndexOf(".xlsx", StringComparison.Ordinal)) + ".xlsx";

            LogDisplay.RecordLine(
                "[{0}] , {1}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                modelName
            );

            // 获取列标题
            var headers = new List<object>();
            foreach (Range cell in list.HeaderRowRange.Cells)
            {
                headers.Add(cell.Value);
            }

            // 获取所有行数据
            var rows = new List<List<object>>();
            foreach (Range row in list.DataBodyRange.Rows)
            {
                var rowData = new List<object>();
                foreach (Range cell in row.Cells)
                {
                    var cellValue = cell.Value?.ToString();

                    foreach (var exportWildcard in exportWildcardData)
                    {
                        if (cellValue != null)
                            cellValue = cellValue.Replace(
                                $"#{exportWildcard.Key}#",
                                exportWildcard.Value
                            );
                    }

                    rowData.Add(cellValue);
                }
                rows.Add(rowData);
            }
            modelValueAll[modelName] = (headers, rows);
        }
        //写入数据
        var excelData = new ExcelDataAutoInsertNumChanges();
        excelData.SetNumChangesData(modelValueAll);
    }
}

//使用Interop获取激活Sheet数据，Epplus写入目标工作簿数据
public static class ExcelDataAutoInsertCopyActivity
{
    public static void RightClickCloneData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var wkPath = NumDesAddIn.App.ActiveWorkbook.Path;
        var excelNames = new List<string>()
        {
            "RechargeAmazon.xlsx",
            "RechargeAptoide.xlsx",
            "RechargeGlobalOfficial.xlsx",
            "RechargeSamsung.xlsx",
            "RechargeIOS.xlsx"
        };
        var defaultValues = new Dictionary<string, List<string>>()
        {
            {"thirdProductID", new List<string>(){ "com.mergeland.alices.adventure_diamond_", "price_Num"}},
        };

        var replaceValues = new Dictionary<string, Dictionary<string, List<string>>>()
        {
            {
                "RechargeIOS.xlsx",
                new Dictionary<string, List<string>>()  
                {
                    {
                        "productID",  
                        new List<string>()
                        {
                            "mergeland.alices.adventure",
                            "casualgame.type.pipe"
                        }
                    },
                    {
                        "productID_test",
                        new List<string>()
                        {
                            "mergeland.alices.adventure",
                            "casualgame.type.pipe"
                        }
                    }
                }  
            }  
        };

        ExcelDataSyncHelper.SyncSelectedRows(
            targetPath: wkPath,
            targetFileNames: excelNames,
            defaultValues: defaultValues,
            replaceValues: replaceValues
            );
    }

    public static void RightClickCloneAllData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var wkPath = NumDesAddIn.App.ActiveWorkbook.Path;
        var excelNames = new List<string>()
        {
            "RechargeAmazon.xlsx",
            "RechargeAptoide.xlsx",
            "RechargeGlobalOfficial.xlsx",
            "RechargeSamsung.xlsx",
            "RechargeIOS.xlsx"
        };
        var defaultValues = new Dictionary<string, List<string>>()
        {
            {"thirdProductID", new List<string>(){ "com.mergeland.alices.adventure_diamond_", "price_Num"}},
        };

        var replaceValues = new Dictionary<string, Dictionary<string, List<string>>>()
        {
            {
                "RechargeIOS.xlsx",
                new Dictionary<string, List<string>>()
                {
                    {
                        "productID",
                        new List<string>()
                        {
                            "mergeland.alices.adventure",
                            "casualgame.type.pipe"
                        }
                    },
                    {
                        "productID_test",
                        new List<string>()
                        {
                            "mergeland.alices.adventure",
                            "casualgame.type.pipe"
                        }
                    }
                }
            }
        };

        ExcelDataSyncHelper.SyncAllRows(
            targetPath: wkPath,
            targetFileNames: excelNames,
            sourceFileName: "RechargeGP.xlsx",
            defaultValues: defaultValues,
            replaceValues: replaceValues
        );

    }

 
}

public static class ExcelDataSyncHelper
{
    private static dynamic ExcelApp => NumDesAddIn.App;

    #region 公共接口方法

    /// <summary>
    /// 同步选中行数据到多个目标工作表
    /// </summary>
    public static void SyncSelectedRows(string targetPath, List<string> targetFileNames,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues)
    {
        NumDesAddIn.App.StatusBar = false;
        var timer = Stopwatch.StartNew();

        try
        {
            // 获取选中数据
            var selection = ExcelApp.Selection as Range;
            if (selection == null) return;

            var sourceData = GetSelectedData(selection);
            if (sourceData.Count == 0) return;

            // 同步到每个目标文件
            foreach (var fileName in targetFileNames)
            {
                SyncDataToTargetFile(targetPath, fileName, sourceData, defaultValues, replaceValues);
                ExcelApp.StatusBar = $"导出：{targetPath}\\{fileName}";
            }

            ExcelApp.StatusBar = $"同步完成，用时：{timer.ElapsedMilliseconds}";
        }
        finally
        {
            timer.Stop();
        }
    }

    /// <summary>
    /// 全量同步源表数据到多个目标工作表(跳过已存在的数据)
    /// </summary>
    public static void SyncAllRows(string targetPath, List<string> targetFileNames,
        string sourceFileName,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues)
    {
        NumDesAddIn.App.StatusBar = false;
        var timer = Stopwatch.StartNew();

        try
        {
            // 读取源表所有数据
            var sourceData = ReadExcelData(targetPath, sourceFileName);
            if (sourceData.Count == 0) return;

            // 同步到每个目标文件
            foreach (var fileName in targetFileNames)
            {
                SyncAllDataToTargetFile(targetPath, fileName, sourceData, defaultValues, replaceValues);
                ExcelApp.StatusBar = $"导出：{targetPath}\\{fileName}";
            }

            ExcelApp.StatusBar = $"全量同步完成，用时：{timer.ElapsedMilliseconds}";
        }
        finally
        {
            timer.Stop();
        }
    }

    #endregion

    #region 核心同步逻辑

    private static void SyncDataToTargetFile(string path, string fileName,
        List<Dictionary<string, object>> sourceData,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues)
    {
        using var package = new ExcelPackage(new FileInfo(Path.Combine(path, fileName)));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
        if (worksheet == null) return;

        var headers = GetHeaders(worksheet);
        int startRow = worksheet.Dimension?.End.Row + 1 ?? 1;

        foreach (var rowData in sourceData)
        {
            WriteRowData(worksheet, headers, startRow++, rowData, defaultValues,
                replaceValues.GetValueOrDefault(fileName));
        }

        package.Save();
    }

    private static void SyncAllDataToTargetFile(string path, string fileName,
        List<Dictionary<string, object>> sourceData,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, Dictionary<string, List<string>>> replaceValues)
    {
        using var package = new ExcelPackage(new FileInfo(Path.Combine(path, fileName)));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
        if (worksheet == null) return;

        var headers = GetHeaders(worksheet);
        var idColumnIndex = headers.IndexOf("id") + 1; // 假设id列存在

        // 获取目标表中已有的ID集合
        var existingIds = new HashSet<string>();
        if (worksheet.Dimension != null && idColumnIndex > 0)
        {
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var idValue = worksheet.Cells[row, idColumnIndex].Text;
                if (!string.IsNullOrEmpty(idValue))
                {
                    existingIds.Add(idValue);
                }
            }
        }

        // 写入新数据
        int startRow = worksheet.Dimension?.End.Row + 1 ?? 1;

        foreach (var rowData in sourceData)
        {
            if (rowData.TryGetValue("id", out var idObj) && idObj != null)
            {
                string id = idObj.ToString();
                if (!existingIds.Contains(id))
                {
                    WriteRowData(worksheet, headers, startRow++, rowData, defaultValues,
                        replaceValues.GetValueOrDefault(fileName));
                }
            }
        }

        package.Save();
    }

    private static void WriteRowData(ExcelWorksheet worksheet, List<string> headers, int row,
        Dictionary<string, object> sourceRow,
        Dictionary<string, List<string>> defaultValues,
        Dictionary<string, List<string>> replaceRules)
    {
        for (int col = 0; col < headers.Count; col++)
        {
            var header = headers[col];
            var cell = worksheet.Cells[row, col + 1];

            if (sourceRow.TryGetValue(header, out var value))
            {
                // 应用替换规则
                if (replaceRules != null && replaceRules.TryGetValue(header, out var replacePair)
                    && replacePair.Count >= 2)
                {
                    cell.Value = value?.ToString()?.Replace(replacePair[0], replacePair[1]);
                }
                else
                {
                    cell.Value = value;
                }
            }
            else if (defaultValues.TryGetValue(header, out var defaultValue)
                     && defaultValue.Count >= 2
                     && sourceRow.TryGetValue(defaultValue[1], out var refValue))
            {
                // 应用默认值公式
                cell.Value = defaultValue[0] + Convert.ToDouble(refValue) / 100;
            }
        }
    }

    #endregion

    #region 数据读取方法

    private static List<Dictionary<string, object>> GetSelectedData(Range selection)
    {
        var headers = GetHeaders(ExcelApp.ActiveSheet);
        var data = new List<Dictionary<string, object>>();

        foreach (Range row in selection.Rows)
        {
            var rowData = new Dictionary<string, object>();

            for (int i = 0; i < headers.Count; i++)
            {
                rowData[headers[i]] = row.Cells[1, i + 1].Value;
            }

            data.Add(rowData);
        }

        return data;
    }

    private static List<Dictionary<string, object>> ReadExcelData(string path, string fileName)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var filePath = Path.Combine(path, fileName);
        if (!File.Exists(filePath)) return new List<Dictionary<string, object>>();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
        return worksheet == null
            ? new List<Dictionary<string, object>>()
            : ExcelToDictionaryList(worksheet);
    }

    private static List<Dictionary<string, object>> ExcelToDictionaryList(ExcelWorksheet worksheet, int headerRow = 2)
    {
        var data = new List<Dictionary<string, object>>();
        if (worksheet.Dimension == null) return data;

        // 读取表头
        var headers = new Dictionary<int, string>();
        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
        {
            headers[col] = worksheet.Cells[headerRow, col].Text ?? $"Column{col}";
        }

        // 读取数据行
        for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
        {
            var rowData = new Dictionary<string, object>();
            bool hasData = false;

            foreach (var header in headers)
            {
                var value = worksheet.Cells[row, header.Key].Value;
                if (value != null) hasData = true;
                rowData[header.Value] = value;
            }

            if (hasData) data.Add(rowData);
        }

        return data;
    }

    #endregion

    #region 辅助方法

    private static List<string> GetHeaders(dynamic sheet)
    {
        if (sheet is ExcelWorksheet eppSheet)
        {
            return GetEpplusHeaders(eppSheet);
        }
        else if (sheet is Worksheet comSheet)
        {
            return GetComHeaders(comSheet);
        }
        return new List<string>();
    }

    private static List<string> GetComHeaders(Worksheet worksheet)
    {
        try
        {
            var rowRange = worksheet.Rows[2] as Range;
            var values = rowRange?.Value as object[,];

            return values == null
                ? new List<string>()
                : Enumerable.Range(1, values.GetLength(1))
                    .Select(col => values[1, col]?.ToString() ?? "")
                    .ToList();
        }
        catch
        {
            return new List<string>();
        }
    }

    private static List<string> GetEpplusHeaders(ExcelWorksheet worksheet)
    {
        if (worksheet.Dimension == null) return new List<string>();

        return Enumerable.Range(worksheet.Dimension.Start.Column, worksheet.Dimension.End.Column)
            .Select(col => worksheet.Cells[2, col].Text ?? "")
            .ToList();
    }

    #endregion
}