using OfficeOpenXml;

#pragma warning disable CA1416

using NumDesTools;
using NumDesTools.AutoInsert;
using NumDesTools.Export;

namespace NumDesTools.AutoInsert;

public static class ExcelDataAutoInsertCopyMulti
{
    public static void SearchData(dynamic isMulti)
    {
        var indexWk = AppServices.App.ActiveWorkbook;
        var sheet = AppServices.App.ActiveSheet;
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
                excelPath != AppServices.Config.Paths.TempPath
                    ? AppServices.Config.Paths.TempPath
                    : AppServices.Config.Paths.BasePath;
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

            AppServices.App.StatusBar =
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
                tempWorkbook = AppServices.App.Workbooks.Open(
                    excelPath + @"\#合并表格数据缓存.xlsx"
                );
            }
            catch
            {
                tempWorkbook = AppServices.App.Workbooks.Add();
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
            AppServices.App.Visible = true;
            AppServices.App.StatusBar = "完成统计";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }
}
