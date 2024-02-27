using ExcelDna.Integration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Color = System.Drawing.Color;
using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Match = System.Text.RegularExpressions.Match;
using MessageBox = System.Windows.Forms.MessageBox;
using Range = Microsoft.Office.Interop.Excel.Range;
#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// Merge项目Excel文件数据自动处理类集合
/// </summary>
public class ExcelDataAutoInsert
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

    public static int FindSourceCol(ExcelWorksheet sheet, int row, string searchValue)
    {
        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
        {
            // 获取当前行的单元格数据
            var cellValue = sheet.Cells[row, col].Value;

            // 如果找到了匹配的值
            if (cellValue != null && cellValue.ToString() == searchValue)
            {
                // 返回该单元格的行地址
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Column;
                return rowAddress;
            }
        }

        return -1;
    }

    public static int FindSourceRow(ExcelWorksheet sheet, int col, string searchValue)
    {
        for (var row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            // 获取当前行的单元格数据
            var cellValue = sheet.Cells[row, col].Value;

            // 如果找到了匹配的值
            if (cellValue != null && cellValue.ToString() == searchValue)
            {
                // 返回该单元格的行地址
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Row;
                return rowAddress;
            }
        }

        return -1;
    }

    [ExcelFunction(IsHidden = true)]
    public static string ErrorExcelMark(dynamic errorExcelList, dynamic sheet)
    {
        var strBuild = new StringBuilder();
        for (var i = 0; i < errorExcelList.Count; i++)
        {
            if (errorExcelList[i][0].Item1 == 0) continue;
            strBuild.Append(errorExcelList[i][0].Item2);
            var cell = sheet.Cells[errorExcelList[i][0].Item1, 1];
            cell.Value = "git checkout -- Excels/Tables/" + errorExcelList[i][0].Item3;
            cell.Font.Color = Color.Red;
            //cell.Font.Size = 9;
            //cell.Font.Name = "微软雅黑";
            //cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //cell.VerticalAlignment = XlHAlign.xlHAlignCenter;
            //cell.ColumnWidth = 8.38;
            //cell.RowHeight = 14.25;
            //cell.ShrinkToFit = true;
            //cell.Borders.LineStyle = XlLineStyle.xlDash;
            //cell.Borders.Weight = XlBorderWeight.xlHairline;
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
            var num = long.Parse(numStr);
            if (digit.Any(item => item.Item1 == matchCount + 1))
            {
                //数字累加
                var addDigit = (long)Math.Pow(10, digit[digitCount].Item2 - 1) * addValue;
                //if (addDigit >= num)
                //{
                //    str = "^error^";
                //    return str;
                //}
                var newNum = num + addDigit;
                //字符替换
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
                digitCount++;
            }
            else if (digit.Count == 1 && digit[0].Item1 == 0)
            {
                //数字累加
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
                //字符替换
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
            //找到模板表所在行
            var modeCol = FindTitle(sheet, 1, "实际模板(上一期)");
            var excelName = FindTitle(sheet, 1, "表名");
            string findValue = sheet.Cells[i, modeCol].Value?.ToString();
            var cell = sheet.Cells[i, excelName];
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx")) continue;
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
            var row = FindSourceRow(sheetTemp, 2, findValue);
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
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx")) continue;
            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
            string path = cell.value switch
            {
                "Localizations.xlsx" => newPath + @"\Excels\Localizations\Localizations.xlsx",
                "UIConfigs.xlsx" => newPath + @"\Excels\UIs\UIConfigs.xlsx",
                "UIItemConfigs.xlsx" => newPath + @"\Excels\UIs\UIItemConfigs.xlsx",
                _ => excelPath + @"\" + cell.value
            };

            //var excel = new ExcelPackage(new FileInfo(path));
            //var workbook = excel.Workbook;
            //var sheetTemp = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

            //var sheetName = sheetTemp.Name;
            var links = path + "#" + "Sheet1!A1";
            //excel.Dispose();
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
                    if (!int.TryParse(parts[0], out var key)) // 尝试将值解析为整数，如果解析失败就将值设为 0
                    {
                        MessageBox.Show($@"{str}#前必须有数值");
                        Environment.Exit(0);
                    }

                    if (!int.TryParse(parts[1], out var value)) // 尝试将值解析为整数，如果解析失败就将值设为 0
                        value = 1;

                    monkeyList.Add((key, value));
                }
                else
                {
                    monkeyList.Add((int.Parse(pair), 1));
                }
        }
        else
        {
            if (str.Contains('#'))
            {
                var parts = str.Split('#');
                var key = Convert.ToInt32(parts[0]);
                var value = Convert.ToInt32(parts[1]);
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
                    strTemp = int.Parse(str);
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

public class ExcelDataAutoInsertLanguage
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
        if (errorExcelList == null) throw new ArgumentNullException(nameof(errorExcelList));

        List<(int, string, string)> error =
            LanguageDialogData(sourceSheet, fixSheet, classSheet, emoSheet, excelPath, NumDesAddIn.App);

        if (error.Count != 0) errorExcelList.Add(error);

        //出错表格处理
        string errorLog = ExcelDataAutoInsert.ErrorExcelMark(errorExcelList, fixSheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }

        //else
        //{
        //    fixSheet.Range["A2:A1000"].Value = "";
        //}
        //app.Dispose();
        Marshal.ReleaseComObject(sourceSheet);
        Marshal.ReleaseComObject(fixSheet);
        Marshal.ReleaseComObject(classSheet);
        Marshal.ReleaseComObject(emoSheet);
        Marshal.ReleaseComObject(workBook);
    }

    public static List<(int, string, string)> LanguageDialogData(dynamic sourceSheet, dynamic fixSheet,
        dynamic classSheet, dynamic emoSheet, string excelPath, dynamic app)
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
            if (fixDataList[i][fileIndex] == null) continue;
            //整理要修改的表格和原表格数据字段映射
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

            //遍历要修改的表格写入数据
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

            //数据查重
            var c = 0;
            if (fixFileName == "GuideDialogDetail.xlsx")
                c = 1;
            else if (fixFileName == "Localizations.xlsx")
                c = 2;
            else if (fixFileName == "GuideDialogBranch.xlsx") c = 3;
            var idList = new List<string>();
            for (var r = 0; r < sourceDataList.Count; r++)
            {
                var value = sourceDataList[r][c]?.ToString() ?? "";
                idList.Add(value);
            }

            var newIdList = idList.Distinct().ToList();

            // 定义要删除的行的列表
            var rowsToDelete = new List<int>();
            foreach (var id in newIdList)
            {
                var reDd = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, id);
                if (reDd != -1) rowsToDelete.Add(reDd);
            }

            //int endRow = targetSheet.Dimension.End.Row;
            //// 遍历行并找到具有相同第一列值的行
            //for (var row = 4; row <= endRow; row++)
            //{
            //    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";
            //    if (idSet.Contains(cellValue))
            //    {
            //        // 如果发现第一列值相同的行，则删除该行
            //        targetSheet.DeleteRow(row);
            //        // 调整删除后的行号
            //        row--;
            //        endRow--;
            //    }
            //}
            rowsToDelete.Sort();
            rowsToDelete.Reverse();
            // 逐行删除
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

            //根据模板插入对应数据行，并复制
            var endRowSource = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, fixFileModeId);
            if (endRowSource == -1)
            {
                MessageBox.Show(fixFileModeId + @"目标表中不存在");
                continue;
            }

            targetSheet.InsertRow(endRowSource + 1, sourceDataList.Count);
            var colCount = targetSheet.Dimension.Columns;
            var cellSource = targetSheet.Cells[endRowSource, 1, endRowSource, colCount];
            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var cellTarget = targetSheet.Cells[endRowSource + 1 + m, 1, endRowSource + 1 + m, colCount];
                cellSource.Copy(cellTarget,
                    ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting |
                    ExcelRangeCopyOptionFlags.ExcludeMergedCells);
                cellSource.CopyStyles(cellTarget);
            }

            //修改数据
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
                    var cellCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, fixKeyList[sourceCount]);
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
                        else if (fixFileName == "GuideDialogBranch.xlsx") realCol = "BranchID";
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(realCol)];
                        if (sourceValue == "" || sourceValue == null) continue;
                        var str = sourceValue.ToString();
                        var digit = Math.Pow(10, e);
                        var repeatCount = 0;
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf(realCol)];
                            if (repeatValue == "" || repeatValue == null) continue;
                            if (repeatValue == sourceValue)
                            {
                                var newNum = long.Parse(str) * digit + repeatCount + 1;
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
                        if (sourceStr == null || sourceStr == "") continue;
                        var matches = Regex.Matches(sourceStr, reg);

                        var oldId = matches[0].Value.ToString();
                        if (newId != "") sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色表情")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)].ToString();
                        for (var k = 0; k < emoDataList.Count; k++)
                        {
                            var targetValue = emoDataList[k][0].ToString();
                            if (targetValue == sourceValue)
                            {
                                var emoId = emoDataList[k][2];
                                cellTarget.Value = emoId;
                                break;
                            }
                        }
                    }
                    else if (source == "触发分支")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)]?.ToString();
                        if (sourceValue == null || sourceValue == "" || sourceValue == "0")
                        {
                            sourceCount++;
                            continue;
                        }

                        var uniqueValues1 = new HashSet<string>();
                        var strBranch = "";
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf("分支归属")]?.ToString();
                            if (repeatValue == null || repeatValue == "") continue;
                            if (repeatValue == sourceValue)
                            {
                                var branchId = sourceDataList[k][sourceTitle.IndexOf("BranchID")];
                                if (uniqueValues1.Add(branchId))
                                {
                                    strBranch = strBranch + branchId + ",";
                                }
                            }
                        }

                        strBranch = "[" + strBranch.Substring(0, strBranch.Length - 1) + "]";
                        cellTarget.Value = strBranch;
                    }
                    else if (source == "分支多语言")
                    {
                        var newId = sourceDataList[m][sourceTitle.IndexOf("BranchID")]?.ToString();
                        var sourceStr = cellTarget.Value?.ToString();
                        if (sourceStr == null || sourceStr == "") continue;
                        var reg = "\\d+";
                        var matches = Regex.Matches(sourceStr, reg);
                        var oldId = matches[0].Value.ToString();
                        if (newId != "") sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色换装1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m][sourceTitle.IndexOf("角色换装")]?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue = sourceValue2 == "1" ? (string)classDataList[k][scCol + 2].ToString() : "[]";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "角色换装2")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m][sourceTitle.IndexOf("角色换装")]?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue = sourceValue2 != "1" ? (string)classDataList[k][scCol + 3].ToString() : "";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "UI对话框")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("UI对话框")];
                        sourceValue = sourceValue == null ? "1" : sourceValue.ToString();
                        if (fixKeyList[sourceCount] == "bgType") cellTarget.Value = sourceValue;
                    }
                    else
                    {
                        //GroupID不连续
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)];
                        cellTarget.Value = sourceValue;
                    }

                    sourceCount++;
                }
            }

            //数据合并
            if (errorExcel != 0) continue;
            int startRow = endRowSource + 1;
            int endRow2 = startRow + sourceDataList.Count - 1;
            //int endRow2 = targetSheet.Dimension.End.Row;
            //if (hasCopy == 2)
            //{
            //    var dataCount = int.Parse(fixDataList[i][dataRows]);
            //    var start = endRow2 + 1;
            //    targetSheet.DeleteRow(start, dataCount);
            //}
            if (fixFileName == "GuideDialogBranch.xlsx" || fixFileName == "GuideDialogGroup.xlsx")
            {
                var uniqueValues = new HashSet<string>();
                // 遍历行并找到具有相同第一列值的行
                for (var row = 4; row <= endRow2; row++)
                {
                    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";

                    if (uniqueValues.Contains(cellValue) || cellValue == "")
                    {
                        // 如果发现第一列值相同的行，则删除该行
                        targetSheet.DeleteRow(row);
                        // 调整删除后的行号
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
            app.StatusBar = "写入数据" + "<" + excelCount + "/" + fixDataList.Count / 2 + ">" + fixFileName;
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
        if (errorExcelList == null) throw new ArgumentNullException(nameof(errorExcelList));

        List<(int, string, string)> error =
            LanguageDialogDataByUd(sourceSheet, fixSheet, classSheet, emoSheet, excelPath, NumDesAddIn.App);

        if (error.Count != 0) errorExcelList.Add(error);

        //出错表格处理
        string errorLog = ExcelDataAutoInsert.ErrorExcelMark(errorExcelList, fixSheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }

        //else
        //{
        //    fixSheet.Range["A2:A1000"].Value = "";
        //}
        //app.Dispose();
        Marshal.ReleaseComObject(sourceSheet);
        Marshal.ReleaseComObject(fixSheet);
        Marshal.ReleaseComObject(classSheet);
        Marshal.ReleaseComObject(emoSheet);
        Marshal.ReleaseComObject(workBook);
    }

    public static List<(int, string, string)> LanguageDialogDataByUd(dynamic sourceSheet, dynamic fixSheet,
        dynamic classSheet, dynamic emoSheet, string excelPath, dynamic app)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //确定行，不确定列
        var sourceData = PubMetToExcel.ExcelDataToListBySelfToEnd(sourceSheet,0,1,1);
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
            if (fixDataList[i][fileIndex] == null) continue;
            //整理要修改的表格和原表格数据字段映射
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

            //遍历要修改的表格写入数据
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

            //数据查重
            var c = 0;
            if (fixFileName == "GuideDialogDetail.xlsx")
                c = 1;
            else if (fixFileName == "Localizations.xlsx")
                c = 2;
            else if (fixFileName == "GuideDialogBranch.xlsx") c = 3;
            var idList = new List<string>();
            for (var r = 0; r < sourceDataList.Count; r++)
            {
                var value = sourceDataList[r][c]?.ToString() ?? "";
                idList.Add(value);
            }

            var newIdList = idList.Distinct().ToList();

            // 定义要删除的行的列表
            var rowsToDelete = new List<int>();
            foreach (var id in newIdList)
            {
                var reDd = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, id);
                if (reDd != -1) rowsToDelete.Add(reDd);
            }

            //int endRow = targetSheet.Dimension.End.Row;
            //// 遍历行并找到具有相同第一列值的行
            //for (var row = 4; row <= endRow; row++)
            //{
            //    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";
            //    if (idSet.Contains(cellValue))
            //    {
            //        // 如果发现第一列值相同的行，则删除该行
            //        targetSheet.DeleteRow(row);
            //        // 调整删除后的行号
            //        row--;
            //        endRow--;
            //    }
            //}
            rowsToDelete.Sort();
            rowsToDelete.Reverse();
            // 逐行删除
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

            //根据模板插入对应数据行，并复制
            //var endRowSource1 = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, fixFileModeId);
            //if (endRowSource == -1)
            //{
            //    MessageBox.Show(fixFileModeId + @"目标表中不存在");
            //    continue;
            //}

            //统一在最后一行插入
            var endRowSource = targetSheet.Dimension.End.Row;

            targetSheet.InsertRow(endRowSource + 1, sourceDataList.Count);
            var colCount = targetSheet.Dimension.Columns;
            var cellSource = targetSheet.Cells[endRowSource, 1, endRowSource, colCount];
            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var cellTarget = targetSheet.Cells[endRowSource + 1 + m, 1, endRowSource + 1 + m, colCount];
                cellSource.Copy(cellTarget,
                    ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting |
                    ExcelRangeCopyOptionFlags.ExcludeMergedCells);
                cellSource.CopyStyles(cellTarget);
            }

            //修改数据
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
                    var cellCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, fixKeyList[sourceCount]);
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
                        else if (fixFileName == "GuideDialogBranch.xlsx") realCol = "BranchID";
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(realCol)];
                        if (sourceValue == "" || sourceValue == null) continue;
                        var str = sourceValue.ToString();
                        var digit = Math.Pow(10, e);
                        var repeatCount = 0;
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf(realCol)];
                            if (repeatValue == "" || repeatValue == null) continue;
                            if (repeatValue == sourceValue)
                            {
                                var newNum = long.Parse(str) * digit + repeatCount + 1;
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
                        if (string.IsNullOrEmpty(sourceStr)) continue;
                        var matches = Regex.Matches(sourceStr, reg);

                        var oldId = matches[0].Value;
                        if (newId != "") sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色表情")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)].ToString();
                        for (var k = 0; k < emoDataList.Count; k++)
                        {
                            var targetValue = emoDataList[k][0].ToString();
                            if (targetValue == sourceValue)
                            {
                                var emoId = emoDataList[k][2];
                                cellTarget.Value = emoId;
                                break;
                            }
                        }
                    }
                    else if (source == "触发分支")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)]?.ToString();
                        if (sourceValue == null || sourceValue == "" || sourceValue == "0")
                        {
                            sourceCount++;
                            continue;
                        }

                        var uniqueValues1 = new HashSet<string>();
                        var strBranch = "";
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf("分支归属")]?.ToString();
                            if (repeatValue == null || repeatValue == "") continue;
                            if (repeatValue == sourceValue)
                            {
                                var branchId = sourceDataList[k][sourceTitle.IndexOf("BranchID")];
                                if (uniqueValues1.Add(branchId))
                                {
                                    strBranch = strBranch + branchId + ",";
                                }
                            }
                        }

                        strBranch = "[" + strBranch.Substring(0, strBranch.Length - 1) + "]";
                        cellTarget.Value = strBranch;
                    }
                    else if (source == "分支多语言")
                    {
                        var newId = sourceDataList[m][sourceTitle.IndexOf("BranchID")]?.ToString();
                        var sourceStr = cellTarget.Value?.ToString();
                        if (string.IsNullOrEmpty(sourceStr)) continue;
                        var reg = "\\d+";
                        var matches = Regex.Matches(sourceStr, reg);
                        var oldId = matches[0].Value;
                        if (newId != "") sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色换装1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m][sourceTitle.IndexOf("角色换装")]?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue = sourceValue2 == "1" ? (string)classDataList[k][scCol + 2].ToString() : "[]";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "角色换装2")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m][sourceTitle.IndexOf("角色换装")]?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newValue = sourceValue2 != "1" ? (string)classDataList[k][scCol + 3].ToString() : "";
                                break;
                            }
                        }

                        cellTarget.Value = newValue;
                    }
                    else if (source == "UI对话框")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("UI对话框")];
                        sourceValue = sourceValue == null ? "1" : sourceValue.ToString();
                        if (fixKeyList[sourceCount] == "bgType") cellTarget.Value = sourceValue;
                    }
                    else
                    {
                        //GroupID不连续
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)];
                        cellTarget.Value = sourceValue;
                    }

                    sourceCount++;
                }
            }

            //数据合并
            if (errorExcel != 0) continue;
            int startRow = endRowSource + 1;
            int endRow2 = startRow + sourceDataList.Count - 1;
            //int endRow2 = targetSheet.Dimension.End.Row;
            //if (hasCopy == 2)
            //{
            //    var dataCount = int.Parse(fixDataList[i][dataRows]);
            //    var start = endRow2 + 1;
            //    targetSheet.DeleteRow(start, dataCount);
            //}
            if (fixFileName == "GuideDialogBranch.xlsx" || fixFileName == "GuideDialogGroup.xlsx")
            {
                var uniqueValues = new HashSet<string>();
                // 遍历行并找到具有相同第一列值的行
                for (var row = 4; row <= endRow2; row++)
                {
                    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";

                    if (uniqueValues.Contains(cellValue) || cellValue == "")
                    {
                        // 如果发现第一列值相同的行，则删除该行
                        targetSheet.DeleteRow(row);
                        // 调整删除后的行号
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
            app.StatusBar = "写入数据" + "<" + excelCount + "/" + fixDataList.Count / 2 + ">" + fixFileName;
        }

        return errorList;
    }
}

public class ExcelDataAutoInsertMulti
{
    public static void InsertData(dynamic isMulti)
    {
        //dynamic app = ExcelDnaUtil.Application;
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
        //var cellBackColor = data[4][baseIdCol];
        var writeMode = data[2][creatIdCol];
        ErrorLogCtp.DisposeCtp();
        //获取单元格颜色
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        //ID自增跨度
        var addValue = (int)data[0][creatIdCol] - (int)data[0][baseIdCol];
        //字典Value跨度（行）
        var rowCount = 2;
        //获取字典
        var colFixKeyCount = colsCount - fixKeyCol;
        var modelId = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdCol, rowCount);
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdNewCol, rowCount);
        var fixKey = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, fixKeyCol, rowCount, colFixKeyCount);
        var ignoreExcel = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, creatIdCol, rowCount);
        //遍历文件写入
        var errorExcelList = new List<List<(string, string, string)>>();
        var excelCount = 1;
        foreach (var key in modelId)
        {
            //写入算法
            var excelName = key.Key;
            //过滤不导出的表格
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

            List<(string, string, string)> error =
                ExcelDataWrite(modelId, modelIdNew, fixKey, excelPath, excelName, addValue, isMulti, commentValue,
                    cellColor, writeMode);
            NumDesAddIn.App.StatusBar = "写入数据" + "<" + excelCount + "/" + modelId.Count + ">" + excelName;
            errorExcelList.Add(error);
            excelCount++;
        }

        //错误日志处理
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //sheet.Range["B4"].Value = "否";
            NumDesAddIn.App.StatusBar = "完成写入";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static void RightClickInsertData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sw = new Stopwatch();
        sw.Start();

        //dynamic app = ExcelDnaUtil.Application;
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
        //var cellBackColor = data[4][baseIdCol];
        var writeMode = data[2][creatIdCol];
        ErrorLogCtp.DisposeCtp();
        //获取单元格颜色
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        //ID自增跨度
        var addValue = (int)data[0][creatIdCol] - (int)data[0][baseIdCol];
        //字典Value跨度（行）
        var rowCount = 2;
        //获取字典
        var colFixKeyCount = colsCount - fixKeyCol;
        var modelId = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdCol, rowCount);
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdNewCol, rowCount);
        var fixKey = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, fixKeyCol, rowCount, colFixKeyCount);
        //遍历文件写入
        var errorExcelList = new List<List<(string, string, string)>>();
        var cell = NumDesAddIn.App.Selection;
        var rowStart = cell.Row;
        var rowCountNew = cell.Rows.Count;
        var rowEnd = rowStart + rowCountNew - 1;
        var excelList = new List<string>();
        //获得一导出文件集合
        for (int i = rowStart; i <= rowEnd; i++)
        {
            var excelName = data[i - 2][sheetNameCol];
            excelList.Add(excelName);
        }

        //去重
        var newExcelList = excelList.Where(excelName => !string.IsNullOrEmpty(excelName)).Distinct().ToList();
        for (var i = 0; i < newExcelList.Count; i++)
        {
            //写入算法
            var excelName = newExcelList[i];
            if (excelName == null) continue;
            List<(string, string, string)> error =
                ExcelDataWrite(modelId, modelIdNew, fixKey, excelPath, excelName, addValue, false, commentValue,
                    cellColor, writeMode);
            NumDesAddIn.App.StatusBar = "写入数据" + "<" + i + "/" + newExcelList.Count + ">" + excelName;
            errorExcelList.Add(error);
        }

        //错误日志处理
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //sheet.Range["B4"].Value = "否";
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

    public static List<(string, string, string)> ExcelDataWrite(dynamic modelId, dynamic modelIdNew, dynamic fixKey,
        dynamic excelPath, dynamic excelName, dynamic addValue, dynamic modeThread, dynamic commentValue,
        dynamic cellBackColor, dynamic writeMode)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorExcelLog = "";
        var errorList = new List<(string, string, string)>();
        string path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        //兼容多表格的工作簿
        string sheetRealName = "Sheet1";
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
        //检查公式
        foreach (var cell in sheet.Cells)
            // 检查单元格是否包含公式
            if (cell.Formula is { Length: > 0 })
            {
                errorList.Add((excelRealName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
                return errorList;
            }

        //获取要查重的ID
        var writeIdList = ExcelDataWriteIdGroup(excelName, addValue, sheet, fixKey, modelId);
        //执行查重删除
        PubMetToExcel.RepeatValue2(sheet, 4, 2, writeIdList.Item1);
        var colCount = sheet.Dimension.Columns;
        //第一次写入插入行的位置，因为可能删除导致行数变化，需要重新获取一次
        writeIdList = ExcelDataWriteIdGroup(excelName, addValue, sheet, fixKey, modelId);
        var writeRow = writeIdList.Item2;
        if (writeRow == -1)
        {
            errorExcelLog = excelName + "#找不到" + writeIdList.Item1[0];
            errorList.Add((excelName, errorExcelLog, excelName));
            return errorList;
        }

        //执行写入操作
        for (var excelMulti = 0; excelMulti < modelId[excelName].Count; excelMulti++)
        {
            var startValue = modelId[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelId[excelName][excelMulti].Item1[1, 0].ToString();
            //var writeRow = sheet.Dimension.End.Row;

            var startRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, startValue);
            if (startRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
                errorList.Add((startValue, errorExcelLog, excelName));
                return errorList;
            }

            var endRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, endValue);
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

            //数据复制
            var count = endRowSource - startRowSource + 1;
            sheet.InsertRow(writeRow + 1, count);
            var cellSource = sheet.Cells[startRowSource, 1, endRowSource, colCount];
            var cellTarget = sheet.Cells[writeRow + 1, 1, writeRow + count, colCount];
            cellTarget.Value = cellSource.Value;
            //设置背景色
            cellTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cellTarget.Style.Fill.BackgroundColor.SetColor(cellBackColor);

            //尽量和源数据格式一致，宋体的格式会歪？不知道为什么
            cellTarget.Style.Font.Name = "微软雅黑";
            cellTarget.Style.Font.Size = 10;
            cellTarget.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            //数据修改
            var fixItem = fixKey[excelName][excelMulti].Item1;
            errorList = modeThread
                ? (List<(string, string, string)>)MultiWrite(excelName, addValue, fixItem, sheet, count, startRowSource,
                    errorList, commentValue, writeRow)
                : (List<(string, string, string)>)SingleWrite(excelName, addValue, fixItem, sheet, count,
                    startRowSource, errorList, commentValue, writeRow);
            writeRow += count;
        }

        excel.Save();
        excel.Dispose();
        errorList.Add(("-1", errorExcelLog, excelName));
        return errorList;
    }

    private static List<(string, string, string)> SingleWrite(dynamic excelName, dynamic addValue, dynamic fixItem,
        ExcelWorksheet sheet,
        dynamic count, dynamic startRowSource, List<(string, string, string)> errorList,
        dynamic commentValue, int writeRow)
    {
        for (var colMulti = 0; colMulti < fixItem.GetLength(1); colMulti++)
        {
            string excelKey = fixItem[0, colMulti];
            if (excelKey == null) continue;
            var excelFileFixKey = ExcelDataAutoInsert.FindSourceCol(sheet, 2, excelKey);
            if (excelFileFixKey == -1)
            {
                var errorExcelLog = excelName + "#【初始模板】#[" + excelKey + "]未找到(字段出错)";
                errorList.Add((excelKey, errorExcelLog, excelName));
                continue;
            }

            string excelKeyMethod = fixItem[1, colMulti]?.ToString();
            //修改字段字典中的字段值，各自方法不一
            for (var i = 0; i < count; i++)
            {
                var cellSource = sheet.Cells[startRowSource + i, excelFileFixKey];
                var rowId = sheet.Cells[startRowSource + i, 2];
                var cellCol = sheet.Cells[2, excelFileFixKey].Value?.ToString();
                var cellFix = sheet.Cells[writeRow + 1 + i, excelFileFixKey];
                if (cellSource.Value == null) continue;

                if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0") continue;
                //默认备注
                if (cellCol != null && cellCol.Contains("#") && commentValue != null)
                {
                    cellFix.Value = commentValue;
                }
                //特殊字段需要CopyAlice
                else
                {
                    //字段值改写方法
                    var temp1 = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                    //修改字符串
                    var cellFixValue = ExcelDataAutoInsert.StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                    if (cellFixValue == "^error^")
                    {
                        string errorExcelLog = excelName + "#" + rowId.Value + "#【修改模式】#[" + excelKey + "]字段方法写错";
                        errorList.Add((excelKey, errorExcelLog, excelName));
                    }

                    cellFix.Value = double.TryParse(cellFixValue, out double number) ? number : cellFixValue;
                }
            }
        }

        return errorList;
    }

    private static List<(string, string, string)> MultiWrite(dynamic excelName, dynamic addValue, dynamic fixItem,
        ExcelWorksheet sheet,
        dynamic count, dynamic startRowSource, List<(string, string, string)> errorList,
        dynamic commentValue, int writeRow)
    {
        var colCoinMulti = fixItem.GetLength(1);
        var colThreadCount = 8; // 线程数
        int colBatchSize = colCoinMulti / colThreadCount; // 每个线程处理的数据量
        Parallel.For(0, colThreadCount, e =>
        {
            var startRow = e * colBatchSize;
            var endRow = (e + 1) * colBatchSize;
            if (e == colThreadCount - 1) endRow = colCoinMulti; // 最后一个线程处理剩余的数据
            for (var k = startRow; k < endRow; k++)
            {
                //查找字段所在列
                string excelKey = fixItem[0, k];
                if (excelKey == null) continue;
                var excelFileFixKey = ExcelDataAutoInsert.FindSourceCol(sheet, 2, excelKey);
                if (excelFileFixKey == -1)
                {
                    var errorExcelLog = excelName + "#【初始模板】#[" + excelKey + "]未找到(字段出错)";
                    errorList.Add((excelKey, errorExcelLog, excelName));
                    continue;
                }

                string excelKeyMethod = fixItem[1, k]?.ToString();

                var rowThreadCount = 4; // 线程数
                int rowBatchSize = count / rowThreadCount; // 每个线程处理的数据量
                // 并发执行任务
                Parallel.For(0, rowThreadCount, i =>
                {
                    var startCol = i * rowBatchSize;
                    var endCol = (i + 1) * rowBatchSize;
                    if (i == rowThreadCount - 1) endCol = count; // 最后一个线程处理剩余的数据

                    for (var j = startCol; j < endCol; j++)
                    {
                        var cellSource = sheet.Cells[startRowSource + j, excelFileFixKey];
                        var cellCol = sheet.Cells[2, excelFileFixKey].Value?.ToString();
                        var cellFix = sheet.Cells[writeRow + j + 1, excelFileFixKey];
                        var rowId = sheet.Cells[startRowSource + j, 2];
                        if (cellSource.Value == null) continue;

                        if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0") continue;

                        if (cellCol != null && cellCol.Contains("#") && commentValue != null)
                        {
                            cellFix.Value = commentValue;
                        }
                        else
                        {
                            //字段值改写方法
                            var temp1 = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                            //修改字符串
                            var cellFixValue =
                                ExcelDataAutoInsert.StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                            if (cellFixValue == "^error^")
                            {
                                string errorExcelLog =
                                    excelName + "#" + rowId.Value + "#【修改模式】#[" + excelKey + "]字段方法写错";
                                errorList.Add((excelKey, errorExcelLog, excelName));
                            }

                            cellFix.Value = double.TryParse(cellFixValue, out double number) ? number : cellFixValue;
                        }
                    }
                });
            }
        });
        return errorList;
    }

    //[ExcelFunction(IsHidden = true)]
    public static (List<string>, int) ExcelDataWriteIdGroup(dynamic excelName, dynamic addValue, ExcelWorksheet sheet,
        dynamic fixKey, dynamic modelId)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var excelFileFixKey = 2;
        var writeIdList = new List<string>();
        var lastRow = 0;
        for (var excelMulti = 0; excelMulti < modelId[excelName].Count; excelMulti++)
        {
            var startValue = modelId[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelId[excelName][excelMulti].Item1[1, 0].ToString();
            var startRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, startValue);
            var endRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, endValue);
            //查错
            if (startRowSource == -1 || endRowSource == -1)
            {
                var writeIdList2 = new List<string> { startValue + "#" + endValue };
                return (writeIdList2, -1);
            }

            string excelKeyMethod = fixKey[excelName][excelMulti].Item1[1, 0]?.ToString();
            //获取要写入的ID
            var count = endRowSource - startRowSource + 1;
            for (var i = 0; i < count; i++)
            {
                var cellSource = sheet.Cells[startRowSource + i, excelFileFixKey];
                if (cellSource.Value == null) continue;
                if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0") continue;
                //字段值改写方法
                var temp1 = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                //修改字符串
                var cellFixValue = ExcelDataAutoInsert.StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                writeIdList.Add(cellFixValue);
            }

            //获取最后一行
            if (lastRow < endRowSource) lastRow = endRowSource;
        }

        return (writeIdList, lastRow);
    }
}

public class ExcelDataAutoInsertCopyMulti
{
    public static void SearchData(dynamic isMulti)
    {
        //dynamic app = ExcelDnaUtil.Application;
        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        //获取单元格颜色
        var colorCell = sheet.Cells[6, 1];
        PubMetToExcel.GetCellBackgroundColor(colorCell);
        ErrorLogCtp.DisposeCtp();
        //字典Value跨度（行）
        var rowCount = 2;
        //获取字典
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdNewCol, rowCount);
        //遍历文件写入
        var errorExcelList = new List<List<(string, string, string)>>();
        var excelCount = 1;
        //统计写入的内容
        var diffList = new List<(string, string, string)>();
        //数据源路径txt
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        foreach (var key in modelIdNew)
        {
            //写入算法
            var excelName = key.Key;
            var targetExcelPath = excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
            //创建excel对象
            List<(string, string, string)> errorList = PubMetToExcel.SetExcelObjectEpPlus(targetExcelPath, excelName,
                out ExcelWorksheet targetSheet, out ExcelPackage _);
            if (errorList.Count != 0)
            {
                //return errorList;
            }

            errorList = PubMetToExcel.SetExcelObjectEpPlus(excelPath, excelName, out ExcelWorksheet sourceSheet,
                out ExcelPackage _);
            if (errorList.Count != 0)
            {
                //return errorList;
            }

            //执行写入操作
            for (var excelMulti = 0; excelMulti < modelIdNew[excelName].Count; excelMulti++)
            {
                var startValue = modelIdNew[excelName][excelMulti].Item1[0, 0].ToString();
                var endValue = modelIdNew[excelName][excelMulti].Item1[1, 0].ToString();
                var startRowSource = ExcelDataAutoInsert.FindSourceRow(sourceSheet, 2, startValue);
                var endRowSource = ExcelDataAutoInsert.FindSourceRow(sourceSheet, 2, endValue);
                var startRowTarget = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, startValue);
                var endRowTarget = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, endValue);
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
                                resultRow = j.ToString();
                            }
                        }

                        if (resultValue != "") diffList.Add((excelPath + @"\" + excelName, resultRow, resultValue));
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
                                resultRow = j.ToString();
                            }
                        }

                        if (resultValue != "")
                            diffList.Add((targetExcelPath + @"\" + excelName, resultRow, resultValue));
                    }
            }

            NumDesAddIn.App.StatusBar = "遍历表格" + "<" + excelCount + "/" + modelIdNew.Count + ">" + excelName;
            errorExcelList.Add(errorList);
            excelCount++;
        }

        //List去重
        diffList = diffList.Distinct().ToList().ToList();
        //错误日志处理
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //人工查询所需要的数据，可以打开表格，可以删除和手动增加数据，专用表格进行操作
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
            // 清除数据（将单元格内容设置为空字符串）
            Range usedRange = tempSheet.UsedRange;
            usedRange.ClearContents();
            //数据转数组
            var tempDataArray = new string[diffList.Count, 4];
            for (var i = 0; i < diffList.Count; i++)
            {
                tempDataArray[i, 0] = diffList[i].Item1;
                tempDataArray[i, 1] = "Sheet1";
                tempDataArray[i, 2] = "B" + diffList[i].Item2;
                tempDataArray[i, 3] = diffList[i].Item3;
            }

            var tempDataRange = tempSheet.Range[tempSheet.Cells[2, 2],
                tempSheet.Cells[2 + tempDataArray.GetLength(0) - 1, 2 + tempDataArray.GetLength(1) - 1]];
            //数据转range
            tempDataRange.Value = tempDataArray;
            tempWorkbook.Save();
            NumDesAddIn.App.Visible = true;
            //if (!File.Exists(outFilePath))
            //{
            //    // 创建新的文本文件
            //    using (StreamWriter writer = File.CreateText(outFilePath))
            //    {
            //        writer.Write(writeRow);
            //        writer.Close();
            //    }
            //}
            //else
            //{
            //    // 覆写已存在的文本文件
            //    using (StreamWriter reader = new StreamWriter(outFilePath, false))
            //    {
            //        reader.Write(writeRow);
            //        reader.Close();
            //    }
            //}
            //Process.Start(outFilePath);
            NumDesAddIn.App.StatusBar = "完成统计";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static void MergeData(dynamic isMulti)
    {
        //dynamic app = ExcelDnaUtil.Application;
        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        //获取单元格颜色
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        ErrorLogCtp.DisposeCtp();
        //字典Value跨度（行）
        var rowCount = 2;
        //获取字典
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdNewCol, rowCount);
        //遍历文件写入
        var errorExcelList = new List<List<(string, string, string)>>();
        var excelCount = 1;
        foreach (var key in modelIdNew)
        {
            //写入算法
            var excelName = key.Key;
            List<(string, string, string)> error = AutoCopyData(modelIdNew, excelName, excelPath, cellColor);
            NumDesAddIn.App.StatusBar = "写入数据" + "<" + excelCount + "/" + modelIdNew.Count + ">" + excelName;
            errorExcelList.Add(error);
            excelCount++;
        }

        //错误日志处理
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //sheet.Range["B4"].Value = "否";
            NumDesAddIn.App.StatusBar = "完成写入";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    private static List<(string, string, string)> AutoCopyData(dynamic modelIdNew, dynamic excelName, dynamic excelPath,
        dynamic cellColor)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorList = new List<(string, string, string)>();
        var targetExcelPath = "";
        //数据源路径txt
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        //第一行Alice，第二行Cove
        if (mergePathList.Count <= 1)
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            //打开文本文件
            Process.Start(filePath);
            return errorList;
        }

        if (mergePathList[0] == "" || mergePathList[1] == "" || mergePathList[1] == mergePathList[0])
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            //打开文本文件
            Process.Start(filePath);
        }
        else
        {
            targetExcelPath = excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        }

        //获取目标、源表格对象
        if (targetExcelPath == "") return errorList;

        //创建excel对象
        errorList = PubMetToExcel.SetExcelObjectEpPlus(targetExcelPath, excelName, out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel);
        if (errorList.Count != 0) return errorList;
        errorList = PubMetToExcel.SetExcelObjectEpPlus(excelPath, excelName, out ExcelWorksheet sourceSheet,
            out ExcelPackage _);
        if (errorList.Count != 0) return errorList;
        //检查公式
        foreach (var cell in targetSheet.Cells)
        {
            // 检查单元格是否包含公式
            if (cell.Formula is not { Length: > 0 }) continue;
            errorList.Add((excelName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
            return errorList;
        }

        //执行写入操作
        for (var excelMulti = 0; excelMulti < modelIdNew[excelName].Count; excelMulti++)
        {
            var startValue = modelIdNew[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelIdNew[excelName][excelMulti].Item1[1, 0].ToString();
            var startRowSource = ExcelDataAutoInsert.FindSourceRow(sourceSheet, 2, startValue);
            string errorExcelLog;
            if (startRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
                errorList.Add((startValue, errorExcelLog, excelName));
                return errorList;
            }

            var endRowSource = ExcelDataAutoInsert.FindSourceRow(sourceSheet, 2, endValue);
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

            //获取目标、源数据表列，以目标表格的行、列为准复制源数据
            var targetMaxCol = targetSheet.Dimension.Columns;
            var sourceMaxCol = sourceSheet.Dimension.Columns;
            var targetRangeTitle = (object[,])targetSheet.Cells[2, 1, 2, targetMaxCol].Value;
            var sourceRangeTitle = (object[,])sourceSheet.Cells[2, 1, 2, sourceMaxCol].Value;
            var sourceRangeValue = (object[,])sourceSheet.Cells[startRowSource, 1, endRowSource, sourceMaxCol].Value;
            //合并数据
            var targetRowList =
                PubMetToExcel.MergeExcel(sourceRangeValue, targetSheet, targetRangeTitle, sourceRangeTitle);
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < targetRowList.Count; i++)
            {
                var cellTarget = targetSheet.Cells[targetRowList[i], 1, targetRowList[i], targetMaxCol];
                var isColorCell = targetSheet.Cells[targetRowList[i], 2];
                // 有填充色就不重新上色了
                if (isColorCell.Style.Fill.BackgroundColor.Rgb == null)
                {
                    //上色
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
        var sw = new Stopwatch();
        sw.Start();

        //dynamic app = ExcelDnaUtil.Application;
        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var excelName = indexWk.Name;
        ErrorLogCtp.DisposeCtp();
        var errorExcelList = new List<List<(string, string, string)>>();
        List<(string, string, string)> error = AutoCopyDataRight(NumDesAddIn.App, excelPath, excelName, sheet);
        errorExcelList.Add(error);
        //错误日志处理
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //sheet.Range["B4"].Value = "否";
            sw.Stop();
            var ts1 = sw.Elapsed;
            NumDesAddIn.App.StatusBar = "完成写入：" + ts1;
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);

        sw.Stop();
        var ts2 = sw.Elapsed;
        NumDesAddIn.App.StatusBar = "完成写入：" + ts2;
        Marshal.ReleaseComObject(sheet);
        Marshal.ReleaseComObject(indexWk);
    }

    private static List<(string, string, string)> AutoCopyDataRight(dynamic app, dynamic excelPath, dynamic excelName,
        dynamic sheet)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorList = new List<(string, string, string)>();
        //数据源路径txt
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        //第一行Alice，第二行Cove

        if (mergePathList.Count <= 1 || mergePathList[0] == "" || mergePathList[1] == "" ||
            mergePathList[1] == mergePathList[0] || !mergePathList[0].Contains("Tables") ||
            !mergePathList[1].Contains("Tables"))
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            Process.Start(filePath);
            return errorList;
        }

        var targetExcelPath = excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        //获取目标表格对象
        if (targetExcelPath == "") return errorList;

        //创建excel对象
        errorList = PubMetToExcel.SetExcelObjectEpPlus(targetExcelPath, excelName, out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel);
        if (errorList.Count != 0) return errorList;
        errorList = PubMetToExcel.SetExcelObjectEpPlus(excelPath, excelName, out ExcelWorksheet sourceSheet,
            out ExcelPackage _);
        if (errorList.Count != 0) return errorList;
        //检查公式
        foreach (var cell in targetSheet.Cells)
            // 检查单元格是否包含公式
            if (cell.Formula is { Length: > 0 })
            {
                errorList.Add((excelName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
                return errorList;
            }

        //复制source表格数据
        var targetMaxCol = targetSheet.Dimension.Columns;
        var sourceMaxCol = sourceSheet.Dimension.Columns;
        var sourceRangeTitle = sourceSheet.Cells[2, 1, 2, sourceMaxCol];
        var targetRangeTitle = targetSheet.Cells[2, 1, 2, targetMaxCol];
        var selectRange = app.Selection;


        if (selectRange.Cells.Count > 0)
        {
            int minRow = selectRange.Row;
            int maxRow = selectRange.Row + selectRange.Rows.Count - 1;
            var sourceRangeValue = (object[,])sourceSheet.Cells[minRow, 1, maxRow, sourceMaxCol].Value;
            var sourceRangeValueTitle = (object[,])sourceRangeTitle.Value;
            var targetRangeValueTitle = (object[,])targetRangeTitle.Value;
            //获取target字段，匹配source的数据进行过滤，并填充数据
            var targetRowList = PubMetToExcel.MergeExcel(sourceRangeValue, targetSheet, targetRangeValueTitle,
                sourceRangeValueTitle);
            //获取单元格颜色
            var colorCell = sheet.Cells[minRow, 2];
            var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
            // ReSharper disable once ForCanBeConvertedToForeach
            for (var i = 0; i < targetRowList.Count; i++)
            {
                var cellTarget = targetSheet.Cells[targetRowList[i], 1, targetRowList[i], targetMaxCol];
                var isColorCell = targetSheet.Cells[targetRowList[i], 2];
                // 有填充色就不重新上色了
                if (isColorCell.Style.Fill.BackgroundColor.Rgb == null)
                {
                    //上色
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
        var sw = new Stopwatch();
        sw.Start();

        //dynamic app = ExcelDnaUtil.Application;
        var indexWk = NumDesAddIn.App.ActiveWorkbook;
        var sheet = NumDesAddIn.App.ActiveSheet;
        var excelPath = indexWk.Path;
        var excelName = indexWk.Name;
        ErrorLogCtp.DisposeCtp();
        var errorExcelList = new List<List<(string, string, string)>>();
        var error = AutoCopyDataRightCol(NumDesAddIn.App, excelPath, excelName);
        errorExcelList.Add(error);
        //错误日志处理
        var errorLog = PubMetToExcel.ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //sheet.Range["B4"].Value = "否";
            sw.Stop();
            var ts1 = sw.Elapsed;
            NumDesAddIn.App.StatusBar = "完成写入：" + ts1;
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);

        sw.Stop();
        var ts2 = sw.Elapsed;
        NumDesAddIn.App.StatusBar = "完成写入：" + ts2;
    }

    private static List<(string, string, string)> AutoCopyDataRightCol(dynamic app, dynamic excelPath,
        dynamic excelName)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var errorList = new List<(string, string, string)>();
        //数据源路径txt
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        //第一行Alice，第二行Cove

        if (mergePathList.Count <= 1 || mergePathList[0] == "" || mergePathList[1] == "" ||
            mergePathList[1] == mergePathList[0] || !mergePathList[0].Contains("Tables") ||
            !mergePathList[1].Contains("Tables"))
        {
            MessageBox.Show(@"找不到目标表格路径，填写其他工程根目录，1行Alice，2行Cove");
            Process.Start(filePath);
            return errorList;
        }

        var targetExcelPath = excelPath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        //获取目标表格对象
        if (targetExcelPath == "") return errorList;

        //创建excel对象
        errorList = PubMetToExcel.SetExcelObjectEpPlus(targetExcelPath, excelName, out ExcelWorksheet targetSheet,
            out ExcelPackage targetExcel);
        if (errorList.Count != 0) return errorList;
        errorList = PubMetToExcel.SetExcelObjectEpPlus(excelPath, excelName, out ExcelWorksheet sourceSheet,
            out ExcelPackage _);
        if (errorList.Count != 0) return errorList;
        //检查公式
        foreach (var cell in targetSheet.Cells)
            // 检查单元格是否包含公式
            if (cell.Formula is { Length: > 0 })
            {
                errorList.Add((excelName, @"不推荐自动写入，单元格有公式:" + cell.Address, "@@@"));
                return errorList;
            }

        //复制source表格数据
        var targetMaxRow = targetSheet.Dimension.Rows;
        var sourceMaxRow = sourceSheet.Dimension.Rows;
        var sourceRangeTitle = sourceSheet.Cells[2, 2, targetMaxRow, 2];
        var targetRangeTitle = targetSheet.Cells[2, 2, sourceMaxRow, 2];
        var selectRange = app.Selection;


        if (selectRange.Cells.Count > 0)
        {
            int minCol = selectRange.Column;
            int maxCol = selectRange.Column + selectRange.Column.Count - 1;
            var sourceRangeValue = (object[,])sourceSheet.Cells[1, minCol, sourceMaxRow, maxCol].Value;
            var sourceRangeValueTitle = (object[,])sourceRangeTitle.Value;
            var targetRangeValueTitle = (object[,])targetRangeTitle.Value;
            //获取target字段，匹配source的数据进行过滤，并填充数据
            PubMetToExcel.MergeExcelCol(sourceRangeValue, targetSheet, targetRangeValueTitle, sourceRangeValueTitle);

            //for (int i = 0; i < targetRowList.Count; i++)
            //{
            //    //获取单元格颜色
            //    var colorCell = sheet.Cells[i+1, 2];
            //    var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);

            //    var cellTarget = targetSheet.Cells[targetRowList[i], 1, targetRowList[i], targetMaxCol];
            //    var isColorCell = targetSheet.Cells[targetRowList[i], 2];
            //    // 有填充色就不重新上色了
            //    if (isColorCell.Style.Fill.BackgroundColor.Rgb == null)
            //    {
            //        //上色
            //        cellTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        cellTarget.Style.Fill.BackgroundColor.SetColor(cellColor);
            //    }
            //}
        }

        targetExcel.Save();
        targetSheet.Dispose();
        return errorList;
    }
}

public class ExcelDataAutoInsertActivityServer
{
    public static void Source()
    {
        //dynamic app = ExcelDnaUtil.Application;
        var indexWk = NumDesAddIn.App.ActiveWorkbook;

        var sourceSheet = indexWk.Worksheets["运营排期"];
        var targetSheet = indexWk.Worksheets["Sheet1"];
        var fixSheet = indexWk.Worksheets["活动模版"];

        var fixData = PubMetToExcel.ExcelDataToList(fixSheet);
        var fixTitle = fixData.Item1;
        var fixDataList = fixData.Item2;
        var fixNames = fixTitle.IndexOf("活动名称");
        var fixIds = fixTitle.IndexOf("活动id");
        // ReSharper disable once IdentifierTypo
        var fixPushs = fixTitle.IndexOf("前端可获取活动时间");
        var fixPushEnds = fixTitle.IndexOf("停止向前端发送活动时间");
        var fixPreHeats = fixTitle.IndexOf("预热期开始时间");
        var fixOpens = fixTitle.IndexOf("活动开启时间");
        var fixEnds = fixTitle.IndexOf("活动结束时间");
        var fixCloses = fixTitle.IndexOf("活动关闭时间");

        var sourceMaxCol = sourceSheet.UsedRange.Columns.Count;
        var sourceMaxRow = sourceSheet.UsedRange.Rows.Count;
        var sourceRange = sourceSheet.Range[sourceSheet.Cells[5, 3], sourceSheet.Cells[sourceMaxRow, sourceMaxCol]];
        var sourceDateRange = sourceSheet.Range[sourceSheet.Cells[2, 1], sourceSheet.Cells[2, sourceMaxCol]];
        Array sourceDataArr = sourceDateRange.Value2;
        //获取目标活动的时间数据
        var sourceData = new List<(string, double, double, int, int, int)>();
        foreach (var cell in sourceRange)
            // 检查单元格是否是合并单元格
            if (cell.MergeCells)
            {
                // 获取合并的行列范围
                var mergeRange = cell.MergeArea;
                if (cell.Address == mergeRange.Cells[1, 1].Address)
                {
                    var mergeValue = mergeRange.Cells[1, 1].Value2;
                    sourceData.Add((mergeValue, sourceDataArr.GetValue(1, mergeRange.Column),
                        sourceDataArr.GetValue(1, mergeRange.Column + mergeRange.Columns.Count - 1), mergeRange.Row,
                        mergeRange.Column, mergeRange.Column + mergeRange.Columns.Count - 1));
                    //// 打印合并范围的行列信息
                    //Debug.Print($"合并范围：{mergedRange.Address}");
                    // ReSharper disable once CommentTypo
                    //Debug.Print($"合并范围内值：{tempvalue}");
                    //Debug.Print($"合并范围的起始行列：{mergedRange.Row},{mergedRange.Column}");
                    //Debug.Print($"合并范围的结束行列：{mergedRange.Row + mergedRange.Rows.Count - 1},{mergedRange.Column + mergedRange.Columns.Count - 1}");
                }
            }
            else if (cell.Value != null)
            {
                sourceData.Add((cell.Value.ToString(), sourceDataArr.GetValue(1, cell.Column),
                    sourceDataArr.GetValue(1, cell.Column + cell.Columns.Count - 1), cell.Row, cell.Column,
                    cell.Column + cell.Columns.Count - 1));
            }

        //对比活动的索引抓取修正时间数据
        var targetDataList = new List<List<string>>();
        var errorLog = "";
        var unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        for (var j = 0; j < sourceData.Count; j++)
        {
            var exit = false;
            var sourceName = sourceData[j].Item1;
            for (var i = 0; i < fixDataList.Count; i++)
            {
                var fixName = fixDataList[i][fixNames];
                if (fixName != sourceName && !sourceName.Contains("#"))
                {
                }
                else if (sourceName.Contains("#"))
                {
                    exit = true;
                }
                else
                {
                    exit = true;
                    var targetData = new List<string>();
                    var sourceStartTimeLong =
                        (long)(DateTime.FromOADate(sourceData[j].Item2).AddHours(8).ToUniversalTime() - unixEpoch)
                        .TotalSeconds;
                    var sourceEndTimeLong =
                        (long)(DateTime.FromOADate(sourceData[j].Item3).AddHours(8).ToUniversalTime() - unixEpoch)
                        .TotalSeconds;

                    var targetId = fixDataList[i][fixIds];
                    var targetName = sourceName;
                    //时间计算规则
                    string targetPushTimeString = DateTime.FromOADate(sourceData[j].Item2)
                        .AddHours(fixDataList[i][fixPushs] * 24 + 8).ToString();
                    var targetPushTimeLong = sourceStartTimeLong + (long)(fixDataList[i][fixPushs] * 24 * 3600);
                    string targetPushEndTimeString = DateTime.FromOADate(sourceData[j].Item2)
                        .AddHours(fixDataList[i][fixPushEnds] * 24 + 8).ToString();
                    var targetPushEndTimeLong = sourceStartTimeLong + (long)(fixDataList[i][fixPushEnds] * 24 * 3600);
                    string targetPreHeatTimeString = DateTime.FromOADate(sourceData[j].Item2)
                        .AddHours(fixDataList[i][fixPreHeats] * 24 + 8).ToString();
                    var targetPreHeatTimeLong = sourceStartTimeLong + (long)(fixDataList[i][fixPreHeats] * 24 * 3600);
                    string targetOpenTimeString = DateTime.FromOADate(sourceData[j].Item2)
                        .AddHours(fixDataList[i][fixOpens] * 24 + 8).ToString();
                    var targetOpenTimeLong = sourceStartTimeLong + (long)(fixDataList[i][fixOpens] * 24 * 3600);
                    //结束时间开始默认+1，因为表格里记录的是时间点不是段
                    string targetEndTimeString = DateTime.FromOADate(sourceData[j].Item3)
                        .AddHours((fixDataList[i][fixEnds] + 1) * 24 + 8).ToString();
                    var targetEndTimeLong = sourceEndTimeLong + (long)(fixDataList[i][fixEnds] * 24 * 3600);
                    string targetCloseTimeString = DateTime.FromOADate(sourceData[j].Item3)
                        .AddHours((fixDataList[i][fixCloses] + 1) * 24 + 8).ToString();
                    var targetCloseTimeLong = sourceEndTimeLong + (long)(fixDataList[i][fixCloses] * 24 * 3600);
                    targetData.Add(targetId.ToString());
                    targetData.Add(targetName);
                    targetData.Add(targetPushTimeString);
                    targetData.Add(targetPushTimeLong.ToString());
                    targetData.Add(targetPushEndTimeString);
                    targetData.Add(targetPushEndTimeLong.ToString());
                    targetData.Add(targetPreHeatTimeString);
                    targetData.Add(targetPreHeatTimeLong.ToString());
                    targetData.Add(targetOpenTimeString);
                    targetData.Add(targetOpenTimeLong.ToString());
                    targetData.Add(targetEndTimeString);
                    targetData.Add(targetEndTimeLong.ToString());
                    targetData.Add(targetCloseTimeString);
                    targetData.Add(targetCloseTimeLong.ToString());
                    targetDataList.Add(targetData);
                }
            }

            if (exit == false)
                errorLog += "运营排期/" + PubMetToExcel.ChangeExcelColChar(sourceData[j].Item5 - 1) + sourceData[j].Item4 +
                            "\r\n";
        }

        //写入新数据
        var targetStartCol = 2;
        var targetStartRow = 5;
        if (errorLog == "")
        {
            //清除老数据
            var targetRangeOld = targetSheet.Range[targetSheet.Cells[targetStartRow, targetStartCol],
                targetSheet.Cells[targetSheet.UsedRange.Rows.Count, targetSheet.UsedRange.Columns.Count]];
            targetRangeOld.Value = null;

            var rows = targetDataList.Count;
            var columns = targetDataList[0].Count;
            var targetDataArr = new string[rows, columns];
            for (var i = 0; i < rows; i++)
            for (var j = 0; j < columns; j++)
                targetDataArr[i, j] = targetDataList[i][j];
            var targetRange = targetSheet.Range[targetSheet.Cells[targetStartRow, targetStartCol],
                targetSheet.Cells[targetStartRow + targetDataArr.GetLength(0) - 1,
                    targetStartCol + targetDataArr.GetLength(1) - 1]];
            targetRange.Value = targetDataArr;
        }
        else
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtp(errorLog);
            //写入错误日志并提示
            MessageBox.Show(@"有活动找不到，查看错误日志");
        }
    }
}