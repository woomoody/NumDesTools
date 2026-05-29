using System.Text.RegularExpressions;
using OfficeOpenXml;
using Match = System.Text.RegularExpressions.Match;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

public static class ExcelDataAutoInsertLanguage
{
    public static void AutoInsertData()
    {
        var workBook = AppServices.App.ActiveWorkbook;
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
            AppServices.App
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
                        if (string.IsNullOrEmpty(sourceStr))
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
                        if (string.IsNullOrEmpty(sourceValue) || sourceValue == "0")
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
                            if (string.IsNullOrEmpty(repeatValue))
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
        cancelDefault = true; // 阻止默认事件

        var workBook = AppServices.App.ActiveWorkbook;
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
            AppServices.App
        );

        if (error.Count != 0)
            errorExcelList.Add(error);

        string errorLog = ExcelDataAutoInsert.ErrorExcelMark(errorExcelList, fixSheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }

        AppServices.App.StatusBar = "导出完成";
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
                                if (newId == null)
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
                        if (string.IsNullOrEmpty(sourceValue) || sourceValue == "0")
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
                            if (string.IsNullOrEmpty(repeatValue))
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
        cancelDefault = true; // 阻止默认事件
        var workBook = AppServices.App.ActiveWorkbook;
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
            object[,] modelRangeValue = list.Range.Value2;

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
            object[,] modelRangeValue = list.Range.Value2;

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

        AppServices.App.StatusBar = "导出完成";
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
                        LogDisplay.RecordLine($"[{DateTime.Now}] , {fixValue}【角色数据】中不存在");
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
                if (newId != null)
                {
                    var reDd = PubMetToExcel.FindSourceRow(targetSheet, 2, newId.ToString());
                    if (reDd != -1)
                        rowsToDelete.Add(reDd);
                }
            }

            rowsToDelete.Sort();
            rowsToDelete.Reverse();

            foreach (var rowToDelete in rowsToDelete)
                try
                {
                    targetSheet.DeleteRow(rowToDelete, 1);
                }
                catch (Exception e)
                {
                    LogDisplay.RecordLine($"[{DateTime.Now}] , {$"sheet表有问题无法删除: {e.Message}"}");
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
                AppServices.App.StatusBar = $"导出：{fixSheetName}";
            }
            targetExcel?.Dispose();
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
            itemValue = itemValue?.ToString() ?? string.Empty;

            if (!dicValue.ContainsKey((funDy2, itemValue)))
            {
                dicValue[(funDy2, itemValue)] = new List<string>();
            }

            var rawValue = sourceDataList[idCount][sourceTitle.IndexOf(funDy2)];
            string value = rawValue?.ToString() ?? string.Empty;
            dicValue[(funDy2, itemValue)].Add(value);

            if (funDy3 != "0")
            {
                dicValue[(funDy2, itemValue)] = dicValue[(funDy2, itemValue)]
                    .Where(values => !string.IsNullOrEmpty(values)) // 过滤掉 null 和空字符串
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
                LogDisplay.RecordLine($"[{DateTime.Now}] , {findValue}【角色数据】中不存在");
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
