using OfficeOpenXml;
using OfficeOpenXml.Style;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

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

        foreach (var key in modelIdNew)
        {
            var excelName = key.Key;
            var targetExcelPath =
                excelPath != NumDesAddIn.TempPath ? NumDesAddIn.TempPath : NumDesAddIn.BasePath;
            List<(string, string, string)> errorList = PubMetToExcel.SetExcelObjectEpPlus(
                targetExcelPath,
                excelName,
                out ExcelWorksheet targetSheet,
                out ExcelPackage targetExcel
            );
            if (errorList.Count != 0) { }

            errorList = PubMetToExcel.SetExcelObjectEpPlus(
                excelPath,
                excelName,
                out ExcelWorksheet sourceSheet,
                out ExcelPackage sourcExcel
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

            targetExcel?.Dispose();
            sourcExcel?.Dispose();

            NumDesAddIn.App.StatusBar =
                "遍历表格" + "<" + excelCount + "/" + modelIdNew.Count + ">" + excelName;
            errorExcelList.Add(errorList);
            excelCount++;
        }

        diffList = diffList.Distinct().ToList();
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
            out ExcelPackage sourceExcel
        );
        if (errorList.Count != 0)
            return errorList;
        if (PubMetToExcel.ShouldCheckFormula(targetExcelPath, targetSheet.Name))
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

        targetExcel.Dispose();
        sourceExcel?.Dispose();

        return errorList;
    }

    public static void RightClickMergeData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件

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
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);

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
            out ExcelPackage sourcExcel
        );
        if (errorList.Count != 0)
            return errorList;
        if (PubMetToExcel.ShouldCheckFormula(targetExcelPath, targetSheet.Name))
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
        targetExcel.Dispose();
        sourcExcel?.Dispose();

        return errorList;
    }

    public static void RightClickMergeDataCol(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件

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
            out ExcelPackage sourceExcel
        );
        if (errorList.Count != 0)
            return errorList;
        if (PubMetToExcel.ShouldCheckFormula(targetExcelPath, targetSheet.Name))
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
        targetExcel.Dispose();
        sourceExcel?.Dispose();

        return errorList;
    }
}
