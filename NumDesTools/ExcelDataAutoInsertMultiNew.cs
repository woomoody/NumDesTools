using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

#pragma warning disable CA1416

namespace NumDesTools;

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
        _indexWk = AppServices.App.ActiveWorkbook;
        _sheet = AppServices.App.ActiveSheet;
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
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【表名】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【初始模板】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【实际模板(上一期)】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【修改字段】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【模板期号】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【创建期号】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【初始备注】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【当前备注】所在列：{_sheetNameCol}");
        LogDisplay.RecordLine($"[{DateTime.Now}] , 【专属替换】所在列：{_sheetNameCol}");

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
            string excelName = key.Key;
            var ignore = _ignoreExcel[excelName][0].Item1[0, 0];
            if (ignore != null)
            {
                var ignoreStr = ignore.ToString();
                if (ignoreStr == "跳过")
                {
                    AppServices.App.StatusBar = "跳过" + "<" + excelName;
                    excelCount++;
                    continue;
                }
            }

            List<(string, string, string)> error = CopyData(excelName);
            AppServices.App.StatusBar =
                "写入数据" + "<" + excelCount + "/" + _modelId.Count + ">" + excelName;
            _errorExcelList.Add(error);
            excelCount++;
        }

        var errorLog = PubMetToExcel.ErrorLogAnalysis(_errorExcelList, _sheet);
        if (errorLog == "")
        {
            AppServices.App.StatusBar = "完成写入";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static void RightClickInsertDataNew(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件
        InitializeVariables();

        var cell = AppServices.App.Selection;
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
            AppServices.App.StatusBar =
                "写入数据" + "<" + i + "/" + newExcelList.Count + ">" + excelName;
            _errorExcelList.Add(error);
        }

        var errorLog = PubMetToExcel.ErrorLogAnalysis(_errorExcelList, _sheet);
        if (errorLog == "")
        {
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static List<(string, string, string)> CopyData(string excelName)
    {
        var errorExcelLog = "";
        List<(string, string, string)> errorList = PubMetToExcel.SetExcelObjectEpPlus(
            _excelPath,
            excelName,
            out ExcelWorksheet sheet,
            out ExcelPackage excel
        );

        ExcelPackage excelNew = null;

        if (excel == null)
        {
            LogDisplay.RecordLine($"[{DateTime.Now}] , {excelName}不存在，看看是否重命名了");
        }

        if (excel != null)
        {
            var excelRealName = excel.File.Name;

            if (PubMetToExcel.ShouldCheckFormula(excel.File.FullName, sheet.Name))
                foreach (var cell in sheet.Cells)
                    if (cell.Formula is { Length: > 0 })
                    {
                        errorList.Add(
                            (
                                $"{excelRealName}#{sheet.Name}",
                                @"不推荐自动写入，单元格有公式:" + cell.Address,
                                "@@@"
                            )
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

            var sheetModel = sheet;

            if (writeRow == -1)
            {
                // 原始表没数据需要额外判断是否有分表，分表名字和原始名字十分相似
                (string, string) fileInfor = PubMetToExcel.AliceFilePathFix(_excelPath, excelName);

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
                            _excelPath,
                            newFileName,
                            out ExcelWorksheet sheetNew,
                            out excelNew
                        );

                        writeIdList = GetElementIdGroup(excelName, sheetNew, _modelId);
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

                var startRowSource = PubMetToExcel.FindSourceRow(sheetModel, 2, startValue);
                if (startRowSource == -1)
                {
                    errorExcelLog =
                        excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
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
                    errorExcelLog =
                        excelName + "#【初始模板】#[" + endValue + "]起始、终结ID顺序反了";
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
                cellColorTarget.Style.Fill.BackgroundColor.SetColor(_cellColor);

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
                    specialValue,
                    sheetModel
                );
                writeRow += count;
            }
        }

        if (excel != null)
        {
            excel.Save();
        }

        excel?.Dispose();
        excelNew?.Dispose();

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
        dynamic specialValue,
        ExcelWorksheet wkSheetModel
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
                    // 通用批量替换列没做任何类型判断,regex替换后原样写回,数字列会被固化成字符串。
                    CellValueNormalizer.ApplyTo(replaceCell, cellValue);
                }
            }
            //自定义修改（修改方法）
            else if (cellKey != "")
            {
                var excelFileFixKey = PubMetToExcel.FindSourceCol(wkSheet, 2, cellKey);
                for (var i = 0; i < count; i++)
                {
                    var cellSource = wkSheetModel.Cells[startRowSource + i, excelFileFixKey];
                    var rowId = wkSheetModel.Cells[startRowSource + i, 2];

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
                            excelName
                            + "#"
                            + rowId.Value
                            + "#【修改模式】#["
                            + cellKey
                            + "]字段方法写错";
                        errorList.Add((cellKey, errorExcelLog, excelName));
                    }

                    // 往返校验版归一化：见 ExcelDataAutoInsertMulti.SingleWrite 里同款替换的注释。
                    CellValueNormalizer.ApplyTo(cellFix, cellFixValue);
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
        var lastRow = -1;
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
