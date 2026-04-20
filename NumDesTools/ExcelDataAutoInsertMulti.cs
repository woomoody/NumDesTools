using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

#pragma warning disable CA1416

namespace NumDesTools;

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
        cancelDefault = true; // 阻止默认事件

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
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
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
        var wk = NumDesAddIn.App.ActiveWorkbook;
        var wkPath = wk.Path;

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

        ExcelPackage excelNew = null;
        ExcelWorksheet sheetModel = sheet;

        if (writeRow == -1)
        {
            // 原始表没数据需要额外判断是否有分表，分表名字和原始名字十分相似
            (string, string) fileInfor = PubMetToExcel.AliceFilePathFix(wkPath, excelName);

            var filePath = Path.GetDirectoryName(fileInfor.Item1);
            var fileName = Path.GetFileNameWithoutExtension(excelName);

            var otherFiles = Directory.GetFiles(
                filePath,
                $"{fileName}_*.xlsx",
                SearchOption.TopDirectoryOnly
            );
            if (otherFiles.Length != 0)
            {
                foreach (var newFile in otherFiles)
                {
                    var newFileName = Path.GetFileName(newFile);

                    errorList = PubMetToExcel.SetExcelObjectEpPlus(
                        wkPath,
                        newFileName,
                        out ExcelWorksheet sheetNew,
                        out excelNew
                    );

                    writeIdList = ExcelDataWriteIdGroup(excelName, addValue, sheetNew, fixKey, modelId); ;
                    writeRow = writeIdList.Item2;

                    if (writeRow != -1)
                    {
                        sheetModel = sheetNew;
                        break;
                    }

                    excelNew?.Dispose();

                    errorExcelLog = excelName + "#找不到" + writeIdList.Item1[0];
                    errorList.Add((excelName, errorExcelLog, excelName));

                }
            }
            else
            {
                errorExcelLog = excelName + "#找不到" + writeIdList.Item1[0];
                errorList.Add((excelName, errorExcelLog, excelName));
                return errorList;
            }

        }

        for (var excelMulti = 0; excelMulti < modelId[excelName].Count; excelMulti++)
        {
            var startValue = modelId[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelId[excelName][excelMulti].Item1[1, 0].ToString();

            var startRowSource = PubMetToExcel.FindSourceRow(sheetModel, 2, startValue);
            if (startRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
                errorList.Add((startValue, errorExcelLog, excelName));
                return errorList;
            }

            var endRowSource = PubMetToExcel.FindSourceRow(sheetModel, 2, endValue);
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
            if (excelRealName.Contains("Recharge") || sheetModel != sheet)
            {
                writeRow = sheet.Dimension.End.Row;
            }
            var count = endRowSource - startRowSource + 1;
            sheet.InsertRow(writeRow + 1, count);
            var cellSource = sheetModel.Cells[startRowSource, 1, endRowSource, colCount];
            var cellTarget = sheet.Cells[writeRow + 1, 1, writeRow + count, colCount];

            cellTarget.Value = cellSource.Value;
            cellTarget.Style.Font.Name = "微软雅黑";
            cellTarget.Style.Font.Size = 10;
            cellTarget.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            //只对前3列标色
            var cellColorTarget = sheet.Cells[writeRow + 1, 1, writeRow + count, 3];
            cellColorTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cellColorTarget.Style.Fill.BackgroundColor.SetColor(cellBackColor);

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
                        writeRow,
                        sheetModel
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
                        writeRow,
                        sheetModel
                    );
            writeRow += count;
        }

        excel.Save();
        excel?.Dispose();
        excelNew?.Dispose();

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
        int writeRow,
        ExcelWorksheet sheetModel
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
                var cellSource = sheetModel.Cells[startRowSource + i, excelFileFixKey];
                var rowId = sheetModel.Cells[startRowSource + i, 2];
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
        int writeRow,
        ExcelWorksheet sheetModel
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
                                var cellSource = sheetModel.Cells[startRowSource + j, excelFileFixKey];
                                var cellCol = sheetModel.Cells[2, excelFileFixKey].Value?.ToString();
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
                                                cellValue = Regex.Replace(
                                                    cellValue,
                                                    pattern,
                                                    replaceValue
                                                );
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
