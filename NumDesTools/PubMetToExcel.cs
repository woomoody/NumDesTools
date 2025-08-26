using System.Collections.Concurrent;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NumDesTools.Config;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;
using ExcelReference = ExcelDna.Integration.ExcelReference;

// ReSharper disable All

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类
/// </summary>
public static class PubMetToExcel
{
    private static readonly Workbook Wk = NumDesAddIn.App.ActiveWorkbook;

    #region EPPlus与Excel

    public static List<(string, string, string)> OpenOrCreatExcelByEpPlus(
        string excelFilePath,
        string excelName,
        out ExcelWorksheet sheet,
        out ExcelPackage excel
    )
    {
        sheet = null;
        excel = null;
        var path = excelFilePath + @"\" + excelName + @".xlsx";
        if (!File.Exists(excelFilePath))
            using (var packageCreat = new ExcelPackage())
            {
                var sheetCreat = packageCreat.Workbook.Worksheets.Add("Sheet1");
                var excelFile = new FileInfo(path);
                packageCreat.SaveAs(excelFile);
                sheetCreat.Dispose();
            }

        var errorList = SetExcelObjectEpPlus(
            excelFilePath,
            excelName + @".xlsx",
            out sheet,
            out excel
        );
        return errorList;
    }

    public static List<(string, string, string)> SetExcelObjectEpPlus(
        dynamic excelPath,
        dynamic excelName,
        out ExcelWorksheet sheet,
        out ExcelPackage excel
    )
    {
        sheet = null;
        excel = null;
        string errorExcelLog;
        var errorList = new List<(string, string, string)>();
        string path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        string sheetRealName = "Sheet1";
        string excelRealName = excelName;

        if (excelName.Contains("##"))
        {
            var excelRealNameGroup = excelName.Split("##");
            if (excelRealNameGroup.Length == 3)
            {
                excelRealName = excelRealNameGroup[1];
                sheetRealName = excelRealNameGroup[2];
            }
            else
            {
                excelRealName = excelRealNameGroup[1];
            }
            excelPath = excelPath + @"\" + excelRealNameGroup[0];
        }
        else
        {
            if (excelName.Contains("#"))
            {
                var excelRealNameGroup = excelName.Split("#");
                excelRealName = excelRealNameGroup[0];
                sheetRealName = excelRealNameGroup[1];
            }
            else if (excelName.Contains("Sharp"))
            {
                var excelRealNameGroup = excelName.Split("Sharp");
                excelRealName = excelRealNameGroup[0].ToString().Replace("Dorllar", "$");
                sheetRealName = excelRealNameGroup[1];
            }
        }

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

        excel = new ExcelPackage(new FileInfo(path));
        ExcelWorkbook workBook;
        try
        {
            workBook = excel.Workbook;
        }
        catch (Exception ex)
        {
            errorExcelLog = excelRealName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelRealName, errorExcelLog, excelRealName));

            excel?.Dispose();

            return errorList;
        }

        try
        {
            sheet = workBook.Worksheets[sheetRealName];
        }
        catch (Exception ex)
        {
            errorExcelLog = excelRealName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelRealName, errorExcelLog, excelRealName));

            excel?.Dispose();
            return errorList;
        }

        sheet ??= workBook.Worksheets[0];

        return errorList;
    }

    public static void SetExcelObjectEpPlusNormal(
        string excelPath,
        string excelName,
        string sheetName,
        out ExcelWorksheet sheet,
        out ExcelPackage excel
    )
    {
        sheet = null;
        excel = null;
        excel = new ExcelPackage(new FileInfo(excelPath + @"\" + excelName));
        ExcelWorkbook workBook;
        try
        {
            workBook = excel.Workbook;
            try
            {
                sheet = workBook.Worksheets[sheetName];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                excel?.Dispose();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
            excel?.Dispose();
        }
    }

    public static List<int> MergeExcelCol(
        object[,] sourceRangeValue,
        ExcelWorksheet targetSheet,
        object[,] targetRangeTitle,
        object[,] sourceRangeTitle
    )
    {
        var targetColList = new List<int>();
        var defaultCol = targetSheet.Dimension.End.Column;
        var beforTargetCol = defaultCol;
        for (var c = 0; c < sourceRangeValue.GetLength(1); c++)
        {
            var sourceCol = sourceRangeValue[1, c];
            if (sourceCol == null)
                sourceCol = "";

            var targetCol = FindSourceCol(targetSheet, 2, sourceCol.ToString());
            if (targetCol == -1)
            {
                targetSheet.InsertColumn(beforTargetCol + 1, 1);
                targetCol = beforTargetCol + 1;
            }

            beforTargetCol = targetCol;
            for (var i = 0; i < targetRangeTitle.GetLength(0); i++)
            {
                var targetTitle = targetRangeTitle[i, 0];
                if (targetTitle == null)
                    targetTitle = "";

                for (var j = 0; j < sourceRangeTitle.GetLength(0); j++)
                {
                    var sourceTitle = sourceRangeTitle[j, 0];
                    if (sourceTitle == null)
                        sourceTitle = "";

                    if (targetTitle.ToString() == sourceTitle.ToString())
                    {
                        var sourceValue = sourceRangeValue[c, j];
                        if (sourceValue == null)
                            sourceValue = "";

                        var targetCell = targetSheet.Cells[targetCol, i + 1];
                        targetCell.Value = sourceValue;
                    }
                }
            }

            targetColList.Add(targetCol);
        }

        return targetColList;
    }

    public static List<int> MergeExcel(
        object[,] sourceRangeValue,
        ExcelWorksheet targetSheet,
        object[,] targetRangeTitle,
        object[,] sourceRangeTitle
    )
    {
        var targetRowList = new List<int>();
        var defaultRow = targetSheet.Dimension.End.Row;
        var beforTargetRow = defaultRow;
        for (var r = 0; r < sourceRangeValue.GetLength(0); r++)
        {
            var sourceRow = sourceRangeValue[r, 1];
            if (sourceRow == null)
                sourceRow = "";

            var targetRow = FindSourceRow(targetSheet, 2, sourceRow.ToString());
            if (targetRow == -1)
            {
                targetSheet.InsertRow(beforTargetRow + 1, 1);
                targetRow = beforTargetRow + 1;
            }

            beforTargetRow = targetRow;
            for (var i = 0; i < targetRangeTitle.GetLength(1); i++)
            {
                var targetTitle = targetRangeTitle[0, i];
                if (targetTitle == null)
                    targetTitle = "";

                for (var j = 0; j < sourceRangeTitle.GetLength(1); j++)
                {
                    var sourceTitle = sourceRangeTitle[0, j];
                    if (sourceTitle == null)
                        sourceTitle = "";

                    if (targetTitle.ToString() == sourceTitle.ToString())
                    {
                        var sourceValue = sourceRangeValue[r, j];
                        if (sourceValue == null)
                            sourceValue = "";

                        var targetCell = targetSheet.Cells[targetRow, i + 1];
                        targetCell.Value = sourceValue;
                    }
                }
            }

            targetRowList.Add(targetRow);
        }

        return targetRowList;
    }

    public static int FindSourceCol(ExcelWorksheet sheet, int row, string searchValue)
    {
        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
        {
            var cellValue = sheet.Cells[row, col].Text;

            if (cellValue != null && cellValue == searchValue)
            {
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Column;
                return rowAddress;
            }
        }

        return -1;
    }

    public static int FindSourceRow(ExcelWorksheet sheet, int col, string searchValue)
    {
        // 遍历指定列的所有行，从第2行开始
        for (var row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            var cellValue = sheet.Cells[row, col].Text; // 获取单元格的文本值

            // 如果单元格值不为空且匹配目标值
            if (!string.IsNullOrEmpty(cellValue) && cellValue == searchValue)
            {
                return row; // 直接返回匹配的行号
            }
        }

        // 如果未找到匹配值，返回 -1
        return -1;
    }

    public static int FindSourceRowBlur(ExcelWorksheet sheet, int col, Regex regexValue)
    {
        var searchRange = sheet.Cells[2, col, sheet.Dimension.End.Row, col];
        var lastMatch = searchRange
            .Reverse()
            .FirstOrDefault(c => c.Value != null && regexValue.IsMatch(c.Value.ToString()));

        var rowIndex = -1;
        if (lastMatch != null)
        {
            rowIndex = lastMatch.Start.Row;
        }
        return -1;
    }

    //获取Excel的ListObject数据并转为数组,非当前
    public static Dictionary<string, object[,]> GetExcelListObjects(
        string excelPath,
        string excelName,
        string sheetName
    )
    {
        PubMetToExcel.SetExcelObjectEpPlusNormal(
            excelPath,
            excelName,
            sheetName,
            out ExcelWorksheet sheet,
            out ExcelPackage excel
        );

        var listObjectDataDic = new Dictionary<string, object[,]>();
        foreach (var table in sheet.Tables)
        {
            if (table != null)
            {
                var tableName = table.Name;

                object[,] tableData =
                    sheet
                        .Cells[
                            table.Address.Start.Row,
                            table.Address.Start.Column,
                            table.Address.End.Row,
                            table.Address.End.Column
                        ]
                        .Value as object[,];
                listObjectDataDic[tableName] = tableData;
            }
        }
        return listObjectDataDic;
    }

    //获取指定表的名称表，当前
    public static ListObject GetExcelListObjects(string sheetName, string listName)
    {
        LogDisplay.RecordLine(
            "[{1}][{0}][{2}][{3}]",
            $"获取Excel ListObject: {sheetName} - {listName}",
            DateTime.Now.ToString(CultureInfo.InvariantCulture),
            sheetName,
            listName
        );
        var sheet = Wk.Worksheets[sheetName];
        // 获取ListObject并操作
        try
        {
            ListObject listObj = sheet.ListObjects[listName];
            return listObj;
        }
        catch (Exception e)
        {
            LogDisplay.RecordLine(
                "[{1}][{0}][{2}][{3}]",
                $"获取Excel ListObject: {sheetName} - {listName} 不存在-{e}",
                DateTime.Now.ToString(CultureInfo.InvariantCulture),
                sheetName,
                listName
            );
            throw;
        }
    }
    #endregion

    #region C-API与Excel
    //索引从0开始
    [ExcelFunction(IsHidden = true)]
    public static object[,] ReadExcelDataC(
        string sheetName,
        int rowFirst,
        int rowLast,
        int colFirst,
        int colLast
    )
    {
        var sheet = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);
        var range = new ExcelReference(rowFirst, rowLast, colFirst, colLast, sheet.SheetId);
        var rangeValue = range.GetValue();
        object[,] rangeValues;
        if (rangeValue is object[,] arrayValue)
        {
            rangeValues = arrayValue;
        }
        else
        {
            rangeValues = new object[1, 1];
            rangeValues[0, 0] = rangeValue;
        }

        return rangeValues;
    }

    [ExcelFunction(IsHidden = true)]
    public static void WriteExcelDataC(
        string sheetName,
        int rowFirst,
        int rowLast,
        int colFirst,
        int colLast,
        object[,] rangeValue
    )
    {
        var sheet = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);
        var range = new ExcelReference(rowFirst, rowLast, colFirst, colLast, sheet.SheetId);
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            range.SetValue(rangeValue);
        });
    }

    public static Task<(
        ExcelReference currentRange,
        string sheetName,
        string sheetPath
    )> GetCurrentExcelObjectC()
    {
        var tcs =
            new TaskCompletionSource<(
                ExcelReference currentRange,
                string sheetName,
                string sheetPath
            )>();
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                var currentRange = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
                var sheetName = (string)XlCall.Excel(XlCall.xlfGetDocument, 1);
                var sheetPath = (string)XlCall.Excel(XlCall.xlfGetDocument, 2);
                var result = (currentRange, sheetName, sheetPath);
                tcs.SetResult(result);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });
        return tcs.Task;
    }

    #endregion

    #region Excel界面相关

    [ExcelFunction(IsHidden = true)]
    public static int ExcelRangePixelsX(double targetX)
    {
        var workArea = NumDesAddIn.App.ActiveWindow;
        var targetXPoint = targetX * 1.67;
        var targetXPixels = workArea.PointsToScreenPixelsX((int)targetXPoint);
        return targetXPixels;
    }

    [ExcelFunction(IsHidden = true)]
    public static int ExcelRangePixelsY(double targetY)
    {
        var workArea = NumDesAddIn.App.ActiveWindow;
        var targetYPoint = targetY * 1.67;
        var targetYPixels = workArea.PointsToScreenPixelsY((int)targetYPoint);
        return targetYPixels;
    }

    #endregion

    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToList(
        dynamic workSheet
    )
    {
        Range dataRange = workSheet.UsedRange;
        object[,] rangeValue = dataRange.Value;
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        for (var row = 1; row <= rows; row++)
        {
            var rowList = new List<object>();
            for (var column = 1; column <= columns; column++)
            {
                var value = rangeValue[row, column];
                if (row == 1)
                    sheetHeaderCol.Add(value);
                else
                    rowList.Add(value);
            }

            if (row > 1)
                sheetData.Add(rowList);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    //public static (List<object> sheetHeaderCol, List<List<object>> sheetStrikethrough) ExcelStrikethroughToList(
    //    dynamic workSheet
    //)
    //{
    //    Range dataRange = workSheet.UsedRange;
    //    object[,] rangeValue = dataRange.Value;
    //    var rows = rangeValue.GetLength(0);
    //    var columns = rangeValue.GetLength(1);
    //    var sheetStrikethrough = new List<List<object>>();
    //    var sheetHeaderCol = new List<object>();
    //    for (var row = 1; row <= rows; row++)
    //    {
    //        var strikethroughList = new List<object>();
    //        for (var column = 1; column <= columns; column++)
    //        {
    //            var value = rangeValue[row, column];
    //            var strikethrough = dataRange[row, column].Font.Strikethrough;
    //            if (row == 1)
    //                sheetHeaderCol.Add(value);
    //            else
    //                strikethroughList.Add(strikethrough);
    //        }

    //        if (row > 1)
    //            sheetStrikethrough.Add(strikethroughList);
    //    }

    //    var excelData = (sheetHeaderCol, sheetStrikethrough: sheetStrikethrough);
    //    return excelData;
    //}
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToListBySelf(
        dynamic workSheet,
        int dataRow,
        int dataCol,
        int headerRow,
        int headerCol
    )
    {
        Range dataRange = workSheet.UsedRange;
        object[,] rangeValue = dataRange.Value;
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        for (var row = dataRow; row <= rows; row++)
        {
            var rowList = new List<object>();
            for (var column = dataCol; column <= columns; column++)
            {
                var value = rangeValue[row, column];
                if (row == headerRow)
                    sheetHeaderCol.Add(value);
                else
                    rowList.Add(value);
            }

            if (row > 1)
                sheetData.Add(rowList);
        }

        for (var column = headerCol; column <= columns; column++)
        {
            var value = rangeValue[headerRow, column];
            sheetHeaderCol.Add(value);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    public static (
        List<object> sheetHeaderCol,
        List<List<object>> sheetData
    ) ExcelDataToListBySelfToEnd(dynamic workSheet, int dataRow, int dataCol, int headRow)
    {
        Range selectRange = NumDesAddIn.App.Selection;
        Range usedRange = workSheet.UsedRange;
        int dataRowEnd;
        int dataColEnd;
        //填0为选中单元格作为数据，否则全部
        if (dataRow == 0)
        {
            dataRow = selectRange.Row;
            dataRowEnd = dataRow + selectRange.Rows.Count - 1;
        }
        else
        {
            dataRow = usedRange.Row;
            dataRowEnd = dataRow + usedRange.Rows.Count - 1;
        }

        if (dataCol == 0)
        {
            dataCol = selectRange.Column;
            dataColEnd = dataCol + selectRange.Columns.Count - 1;
        }
        else
        {
            dataCol = usedRange.Column;
            dataColEnd = dataCol + usedRange.Columns.Count - 1;
        }

        Range dataRangeStart = workSheet.Cells[dataRow, dataCol];
        Range dataRangeEnd = workSheet.Cells[dataRowEnd, dataColEnd];
        Range dataRange = workSheet.Range[dataRangeStart, dataRangeEnd];
        Range headRangeStart = workSheet.Cells[headRow, dataCol];
        Range headRangeEnd = workSheet.Cells[headRow, dataColEnd];
        Range headRange = workSheet.Range[headRangeStart, headRangeEnd];
        var excelData = RangeToListByVsto(dataRange, headRange, headRow);
        return excelData;
    }

    public static DataTable ExcelDataToDataTable(string filePath)
    {
        var file = new FileInfo(filePath);
        using (var package = new ExcelPackage(file))
        {
            var dataTable = new DataTable();
            var worksheet = package.Workbook.Worksheets["Sheet1"] ?? package.Workbook.Worksheets[0];
            dataTable.TableName = worksheet.Name;
            for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                dataTable.Columns.Add();

            for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                var dataRow = dataTable.NewRow();
                for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
                    dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }
    }

    public static DataTable ExcelDataToDataTableOleDb(string filePath)
    {
        var connectionString =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
        var sheetName = "Sheet1";
        using (var connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();
                var dataTable = new DataTable();

                var schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable != null)
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        // ReSharper disable ConditionIsAlwaysTrueOrFalse
                        if (row is not null && row["TABLE_NAME"].ToString().Equals("Sheet1"))
                        // ReSharper restore ConditionIsAlwaysTrueOrFalse
                        {
                            sheetName = "Sheet1";
                            break;
                        }

                        sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();
                    }

                using (var command = new OleDbCommand($"SELECT * FROM [{sheetName}]", connection))
                {
                    using (var adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                if (sheetName != null)
                    dataTable.TableName = sheetName;
                return dataTable;
            }
            catch (Exception ex)
            {
                Debug.Print("读取 Excel 表格数据出现异常：" + ex.Message);
                return null;
            }
        }
    }

    public static List<(string, string, int, int, string, string)> FindDataInDataTable(
        string fileFullName,
        dynamic dataTable,
        string findValue
    )
    {
        var findValueList = new List<(string, string, int, int, string, string)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        var sheetName = dataTable.TableName.ToString().Replace("$", "");
        foreach (DataRow row in dataTable.Rows)
        foreach (DataColumn column in dataTable.Columns)
            if (isAll)
            {
                if (row is not null && row[column].ToString().Contains(findValue))
                    findValueList.Add(
                        (
                            fileFullName,
                            sheetName,
                            row.Table.Rows.IndexOf(row) + 2,
                            row.Table.Columns.IndexOf(column) + 1,
                            row[1].ToString(),
                            row[2].ToString()
                        )
                    );
            }
            else
            {
                if (row[column].ToString() == findValue)
                    findValueList.Add(
                        (
                            fileFullName,
                            sheetName,
                            row.Table.Rows.IndexOf(row) + 2,
                            row.Table.Columns.IndexOf(column) + 1,
                            row[1].ToString(),
                            row[2].ToString()
                        )
                    );
            }

        return findValueList;
    }

    public static List<(string, string, int, int, string, string)> FindDataInDataTableKey(
        string fileFullName,
        dynamic dataTable,
        string findValue,
        int key
    )
    {
        var findValueList = new List<(string, string, int, int, string, string)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        var sheetName = dataTable.TableName.ToString().Replace("$", "");
        foreach (DataRow row in dataTable.Rows)
            if (isAll)
            {
                if (row is not null && row[key - 1].ToString().Contains(findValue))
                    findValueList.Add(
                        (
                            fileFullName,
                            sheetName,
                            row.Table.Rows.IndexOf(row) + 2,
                            key,
                            row[1].ToString(),
                            row[2].ToString()
                        )
                    );
            }
            else
            {
                if (row[key - 1].ToString() == findValue)
                    findValueList.Add(
                        (
                            fileFullName,
                            sheetName,
                            row.Table.Rows.IndexOf(row) + 2,
                            key,
                            row[1].ToString(),
                            row[2].ToString()
                        )
                    );
            }

        return findValueList;
    }

    public static string[] PathExcelFileCollect(
        List<string> pathList,
        string fileSuffixName,
        string[] ignoreFileNames
    )
    {
        var files = new string[] { };
        foreach (var path in pathList)
        {
            var file = Directory
                .EnumerateFiles(path, fileSuffixName)
                .Where(file =>
                    !ignoreFileNames.Any(ignore => Path.GetFileName(file).Contains(ignore))
                )
                .ToArray();
            files = files.Concat(file).ToArray();
        }

        return files;
    }

    public static Dictionary<string, List<Tuple<object[,]>>> ExcelDataToDictionary(
        dynamic data,
        dynamic dicKeyCol,
        dynamic dicValueCol,
        int valueRowCount,
        int valueColCount = 1
    )
    {
        var dic = new Dictionary<string, List<Tuple<object[,]>>>();

        for (var i = 0; i < data.Count; i++)
        {
            var value = data[i][dicKeyCol];

            if (value == null || value == string.Empty)
                continue;

            var values = new object[valueRowCount, valueColCount];
            for (var k = 0; k < valueRowCount; k++)
            for (var j = 0; j < valueColCount; j++)
            {
                var valueTemp = data[i + k][dicValueCol + j];
                values[k, j] = valueTemp;
            }

            var tuple = new Tuple<object[,]>(values);
            if (dic.TryGetValue(value, out List<Tuple<object[,]>> value1))
            {
                value1.Add(tuple);
            }
            else
            {
                var list = new List<Tuple<object[,]>> { tuple };
                dic.Add(value, list);
            }
        }

        return dic;
    }

    public static string RepeatValue(ExcelWorksheet sheet, int row, int col, string repeatValue)
    {
        var errorLog = "";
        for (var r = sheet.Dimension.End.Row; r >= row; r--)
        {
            var colA = sheet.Cells[r, col].Value?.ToString();
            if (colA == repeatValue)
                try
                {
                    sheet.DeleteRow(r);
                }
                catch (Exception ex)
                {
                    errorLog += $"Error {repeatValue}: {ex.Message}\n";
                }
        }

        return errorLog;
    }

    public static string RepeatValue2(
        ExcelWorksheet sheet,
        int row,
        int col,
        List<string> repeatValue
    )
    {
        var errorLog = "";
        var sourceValues = sheet
            .Cells[row, col, sheet.Dimension.End.Row, col]
            .Select(c => c.Value.ToString())
            .ToList();

        var indexList = new List<int>();
        foreach (var repeat in repeatValue)
        {
            var rowIndex = sourceValues.FindIndex(c => c == repeat);
            if (rowIndex == -1)
                continue;
            rowIndex += row;
            indexList.Add(rowIndex);
        }

        indexList.Sort();
        if (indexList.Count != 0)
        {
            var outputList = new List<List<int>>();
            var start = indexList[0];

            for (var i = 1; i < indexList.Count; i++)
                if (indexList[i] != indexList[i - 1] + 1)
                {
                    outputList.Add(new List<int>() { start, indexList[i - 1] });
                    start = indexList[i];
                }

            outputList.Add(new List<int>() { start, indexList[indexList.Count - 1] });
            outputList.Reverse();
            foreach (var rowToDelete in outputList)
                try
                {
                    sheet.DeleteRow(rowToDelete[0], rowToDelete[1] - rowToDelete[0] + 1);
                }
                catch (Exception)
                {
                    errorLog += $"Error {sheet.Name}:#行号{rowToDelete}背景格式问题，更改背景色重试\n";
                }
        }

        return errorLog;
    }

    public static (string file, string Name, int cellRow, int cellCol) ErrorKeyFromExcelId(
        string rootPath,
        string errorValue
    )
    {
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
        var mainPath = newPath + @"\Excels\Tables\";
        var files1 = Directory
            .EnumerateFiles(mainPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var langPath = newPath + @"\Excels\Localizations\";
        var files2 = Directory
            .EnumerateFiles(langPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var uiPath = newPath + @"\Excels\UIs\";
        var files3 = Directory
            .EnumerateFiles(uiPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var kelangPath = newPath + @"\Excels\Tables\克朗代克\";
        //此路径有可能不存在
        string[] files4 = null;
        if (Directory.Exists(kelangPath))
        {
            files4 = Directory
                .EnumerateFiles(kelangPath, "*.xlsx")
                .Where(file => !Path.GetFileName(file).Contains("#"))
                .ToArray();
        }

        var files = files1.Concat(files2).Concat(files3).ToArray();

        var currentCount = 0;
        var count = files.Length;
        foreach (var file in files)
        {
            var fileName = Path.GetFileName(file);
            if (fileName.Contains("#"))
                continue;
            using (var package = new ExcelPackage(new FileInfo(file)))
            {
                try
                {
                    var wk = package.Workbook;
                    var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                    for (var col = 2; col <= 2; col++)
                    for (var row = 4; row <= sheet.Dimension.End.Row; row++)
                    {
                        var cellValue = sheet.Cells[row, col].Value;

                        if (cellValue != null && cellValue.ToString() == errorValue)
                        {
                            var cellAddress = new ExcelCellAddress(row, col);
                            var cellCol = cellAddress.Column;
                            var cellRow = cellAddress.Row;
                            var tuple = (file, sheet.Name, cellRow, cellCol);
                            return tuple;
                        }
                    }
                }
                catch
                {
                    // ignored
                }
            }

            currentCount++;
            NumDesAddIn.App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
        }

        var tupleError = ("", "", 0, 0);
        return tupleError;
    }

    public static Color GetCellBackgroundColor(Range cell)
    {
        var color = Color.Empty;

        if (cell.Interior.Color != null)
        {
            object excelColor = cell.Interior.Color;
            if (excelColor is double)
            {
                var colorValue = (double)excelColor;
                var intValue = (int)colorValue;
                var red = intValue & 0xFF;
                var green = (intValue & 0xFF00) >> 8;
                var blue = (intValue & 0xFF0000) >> 16;
                color = Color.FromArgb(red, green, blue);
            }
        }

        return color;
    }

    public static string ChangeExcelColChar(int col)
    {
        var a = col / 26;
        var b = col % 26;

        if (a > 0)
            return ChangeExcelColChar(a - 1) + (char)(b + 65);

        return ((char)(b + 65)).ToString();
    }

    public static List<string> ReadWriteTxt(string filePath)
    {
        var textLineList = new List<string>();
        if (!File.Exists(filePath))
        {
            if (filePath != null)
                using (var writer = File.CreateText(filePath))
                {
                    writer.WriteLine("Alice路径");
                    writer.WriteLine("Cove路径");
                    writer.Close();
                }
        }
        else
        {
            using var reader = new StreamReader(filePath);
            while (reader.ReadLine() is { } line)
                textLineList.Add(line);
        }

        return textLineList;
    }

    public static string ErrorLogAnalysis(
        List<List<(string, string, string)>> errorList,
        Worksheet sheet
    )
    {
        var errorLog = "";
        for (var i = 0; i < errorList.Count; i++)
        for (var j = 0; j < errorList[i].Count; j++)
        {
            var errorCell = errorList[i][j].Item1;
            var errorExcelLog = errorList[i][j].Item2;
            var errorExcelName = errorList[i][j].Item3;
            if (errorCell == "-1")
                continue;
            errorLog =
                errorLog + "【" + errorCell + "】" + errorExcelName + "#" + errorExcelLog + "\r\n";
        }

        return errorLog;
    }

    public static string ConvertToExcelColumn(int columnNumber)
    {
        var columnName = "";

        while (columnNumber > 0)
        {
            var remainder = (columnNumber - 1) % 26;
            columnName = (char)('A' + remainder) + columnName;
            columnNumber = (columnNumber - 1) / 26;
        }

        return columnName;
    }

    public static void OpenExcelAndSelectCell(string filePath, string sheetName, string cellAddress)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                // ReSharper disable LocalizableElement
                MessageBox.Show(@"文件不存在，请检查！");
                // ReSharper restore LocalizableElement
                return;
            }

            NumDesAddIn.App.ScreenUpdating = false;
            var workbook = NumDesAddIn.App.Workbooks.Open(filePath);

            Worksheet worksheet = null;
            try
            {
                // 尝试获取工作表
                worksheet = (Worksheet)workbook.Sheets[sheetName];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // 如果工作表不存在，则选择第一个工作表
                worksheet = (Worksheet)workbook.Sheets[1];
            }

            var regex = new Regex(@"^[A-Za-z]+\d+$");
            var cellAddressDefault = "1";
            if (cellAddress != null)
            {
                MatchCollection matches = Regex.Matches(cellAddress, @"\d+");
                cellAddressDefault = matches[0].ToString();
                var realCellAddress = $"B{cellAddressDefault}:Z{cellAddressDefault}";
                var cellRange = worksheet.Range[realCellAddress];

                NumDesAddIn.App.ScreenUpdating = true;
                worksheet.Activate();
                cellRange.Select();
            }

            NumDesAddIn.App.ScreenUpdating = true;
        }
        // ReSharper disable EmptyGeneralCatchClause
        catch (Exception)
        // ReSharper restore EmptyGeneralCatchClause
        { }

        GC.Collect();
    }

    public static void ListToArrayToRange(
        List<List<object>> targetList,
        dynamic workSheet,
        int startRow,
        int startCol
    )
    {
        var rowCount = targetList.Count;
        var columnCount = 0;
        foreach (var innerList in targetList)
        {
            var currentColumnCount = innerList.Count;
            columnCount = Math.Max(columnCount, currentColumnCount);
        }

        var targetDataArr = new object[rowCount, columnCount];
        for (var i = 0; i < rowCount; i++)
        for (var j = 0; j < targetList[i].Count; j++)
            targetDataArr[i, j] = targetList[i][j];
        var targetRange = workSheet.Range[
            workSheet.Cells[startRow, startCol],
            workSheet.Cells[startRow + rowCount - 1, startCol + columnCount - 1]
        ];
        targetRange.Value = targetDataArr;
    }

    //Alice文件路径修正
    public static (string filePath, string sheetName) AliceFilePathFix(
        string workbookPath,
        string selectSheetName
    )
    {
        workbookPath = Path.GetDirectoryName(workbookPath);

        var isMatch = selectSheetName.Contains(".xls");

        string filePath = String.Empty;
        string sheetName = "Sheet1";
        if (isMatch)
        {
            if (selectSheetName.Contains("#") && !selectSheetName.Contains("##"))
            {
                var excelSplit = selectSheetName.Split("#");
                filePath = workbookPath + @"\Tables\" + excelSplit[0];
                sheetName = excelSplit[1];
            }
            else if (selectSheetName.Contains("##"))
            {
                var excelSplit = selectSheetName.Split("##");
                var sharpCount = excelSplit.Length;
                if (selectSheetName.Contains("克朗代克"))
                {
                    filePath = workbookPath + @"\Tables\" + excelSplit[0] + @"\" + excelSplit[1];
                    sheetName = sharpCount == 3 ? excelSplit[2] : "Sheet1";
                }
                else
                {
                    selectSheetName = workbookPath + @"\Tables\" + excelSplit[0];
                    sheetName = excelSplit[1];
                }
            }
            else
            {
                switch (selectSheetName)
                {
                    case "Localizations.xlsx":
                        filePath = workbookPath + @"\Localizations\Localizations.xlsx";
                        break;
                    case "UIConfigs.xlsx":
                        filePath = workbookPath + @"\UIs\UIConfigs.xlsx";
                        break;
                    case "UIItemConfigs.xlsx":
                        filePath = workbookPath + @"\UIs\UIItemConfigs.xlsx";
                        break;
                    default:
                        filePath = workbookPath + @"\Tables\" + selectSheetName;
                        break;
                }

                sheetName = "Sheet1";
            }
        }

        return (filePath, sheetName);
    }

    //二维数组搜索指定行的数据，返回指定行对应列数据
    public static string FindValueInFirstRow(
        object[,] array,
        string value,
        int findIndex = 0,
        int returnIndex = 1
    )
    {
        // 获取数组的列数
        int columns = array.GetLength(1);
        for (int col = 0; col < columns; col++)
        {
            if (array[findIndex, col]?.ToString() == value)
            {
                return array[returnIndex, col]?.ToString();
            }
        }

        // 如果未找到匹配的值，返回 null
        return string.Empty;
    }

    //Range二维数组List化
    public static List<List<object>> RangeDataToList(object[,] rangeValue)
    {
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        var sheetData = new List<List<object>>();
        for (var row = 1; row <= rows; row++)
        {
            var rowList = new List<object>();
            for (var column = 1; column <= columns; column++)
            {
                var value = rangeValue[row, column];
                rowList.Add(value);
            }

            sheetData.Add(rowList);
        }

        return sheetData;
    }

    //二维数组List化
    public static List<List<object>> Array2DDataToList(object[,] rangeValue)
    {
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        var sheetData = new List<List<object>>();
        for (var row = 0; row < rows; row++)
        {
            var rowList = new List<object>();
            for (var column = 0; column < columns; column++)
            {
                var value = rangeValue[row, column];
                rowList.Add(value);
            }

            sheetData.Add(rowList);
        }

        return sheetData;
    }

    //二维List一维化
    public static List<object> List2DToListRowOrCol(
        List<List<object>> twoDimensionalList,
        bool byRow
    )
    {
        List<object> flattenedList = new List<object>();

        if (byRow)
        {
            foreach (var row in twoDimensionalList)
            {
                flattenedList.AddRange(row);
            }
        }
        else
        {
            int columnCount = twoDimensionalList[0].Count;
            for (int col = 0; col < columnCount; col++)
            {
                foreach (var row in twoDimensionalList)
                {
                    flattenedList.Add(row[col]);
                }
            }
        }

        return flattenedList;
    }

    public static List<int> GenerateUniqueRandomList(int minValue, int maxValue, int baseValue)
    {
        var list = new List<int>();

        for (var i = minValue; i <= maxValue; i++)
            list.Add(i + baseValue);

        var random = new Random();
        var n = list.Count;
        for (var i = n - 1; i > 0; i--)
        {
            var j = random.Next(0, i + 1);
            var temp = list[i];
            list[i] = list[j];
            list[j] = temp;
        }

        return list;
    }

    //二维List转二维数组
    public static object[,] ConvertListToArray(List<List<object>> listOfLists)
    {
        // 获取行数
        var rowCount = listOfLists.Count;

        // 获取最大列数（找出最长的子列表）
        var colCount = listOfLists.Max(innerList => innerList.Count);

        // 初始化二维数组
        var twoDArray = new object[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            for (var j = 0; j < colCount; j++)
            {
                // 如果当前列索引超出子列表长度，补充空值（null 或 ""）
                twoDArray[i, j] = j < innerList.Count ? innerList[j] : null;
            }
        }

        return twoDArray;
    }

    public static object[,] ConvertListToArray(List<List<string>> listOfLists)
    {
        // 获取行数
        var rowCount = listOfLists.Count;

        // 获取最大列数（找出最长的子列表）
        var colCount = listOfLists.Max(innerList => innerList.Count);

        // 初始化二维数组
        var twoDArray = new object[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            for (var j = 0; j < colCount; j++)
            {
                // 如果当前列索引超出子列表长度，补充空值（null 或 ""）
                twoDArray[i, j] = j < innerList.Count ? innerList[j] : null;
            }
        }

        return twoDArray;
    }

    //一维List转二维数组
    public static object[,] ConvertList1ToArray(List<object> listOfLists)
    {
        // 获取行数
        var rowCount = listOfLists.Count;

        // 获取最大列数（找出最长的子列表）
        int colCount = 1;

        // 初始化二维数组
        var twoDArray = new object[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            twoDArray[i, colCount - 1] = innerList;
        }

        return twoDArray;
    }

    public static object[,] ConvertList1ToArray(List<string> listOfLists)
    {
        // 获取行数
        var rowCount = listOfLists.Count;

        // 获取最大列数（找出最长的子列表）
        int colCount = 1;

        // 初始化二维数组
        var twoDArray = new string[rowCount, colCount];

        // 遍历每个子列表
        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            twoDArray[i, colCount - 1] = innerList;
        }

        return twoDArray;
    }

    public static string[,] ConvertListArrayToTwoArray(List<string[]> listArray)
    {
        var rowmax = listArray.Count;
        var colmax = listArray[0].GetLength(0);

        // 初始化二维数组
        var arrObjects = new string[rowmax, colmax];

        for (int i = 0; i < rowmax; i++)
        {
            var list = listArray[i];
            for (int j = 0; j < colmax; j++)
            {
                arrObjects[i, j] = list[j];
            }
        }
        return arrObjects;
    }

    //一维List转一维数组
    public static object[] ConvertListToArray(List<object> listOfLists)
    {
        var rowCount = listOfLists.Count;
        var twoDArray = new object[rowCount];

        for (var i = 0; i < rowCount; i++)
        {
            twoDArray[i] = listOfLists[i];
        }

        return twoDArray;
    }

    public static object[] ConvertListToArray(List<string> listOfLists)
    {
        var rowCount = listOfLists.Count;
        var twoDArray = new object[rowCount];

        for (var i = 0; i < rowCount; i++)
        {
            twoDArray[i] = listOfLists[i];
        }

        return twoDArray;
    }

    public static (int row, int column) FindValueInRangeByVsto(
        Range searchRange,
        object valueToFind
    )
    {
        Range foundRange = searchRange.Find(valueToFind);
        if (foundRange != null)
        {
            return (foundRange.Row, foundRange.Column);
        }
        else
        {
            return (-1, -1);
        }
    }

    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) RangeToListByVsto(
        Range rangeData,
        Range rangeHeader,
        int headRow
    )
    {
        object[,] rangeValue = rangeData.Value;
        object[,] headRangeValue = rangeHeader.Value;
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        for (var row = 1; row <= rangeValue.GetLength(0); row++)
        {
            var rowList = new List<object>();
            for (var column = 1; column <= rangeValue.GetLength(1); column++)
            {
                var valueData = rangeValue[row, column];
                rowList.Add(valueData);
            }

            sheetData.Add(rowList);
        }

        for (var column = 1; column <= rangeValue.GetLength(1); column++)
        {
            var value = headRangeValue[headRow, column];
            sheetHeaderCol.Add(value);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    //随机数列表唯一方案
    public static List<List<int>> UniqueRandomMethod(
        int numberOfRolls,
        int numberOfSchemes,
        int maxRand
    )
    {
        var result = new List<List<int>>();
        var seenSchemes = new HashSet<string>();
        var random = new Random();

        for (var i = 0; i < numberOfSchemes; i++)
        {
            var scheme = new List<int>();

            for (var j = 0; j < numberOfRolls; j++)
            {
                var randomNumber = random.Next(1, maxRand + 1);
                scheme.Add(randomNumber);
            }

            var schemeString = string.Join(",", scheme);
            if (seenSchemes.Add(schemeString))
                result.Add([.. scheme]);
        }

        return result;
    }

    //二维数组字典化
    public static Dictionary<int, List<object>> TwoDArrayToDictionary(object[,] array)
    {
        Dictionary<int, List<object>> dictionary = new Dictionary<int, List<object>>();

        int rows = array.GetLength(0);
        int cols = array.GetLength(1);

        for (int i = 0; i < rows; i++)
        {
            List<object> rowArray = new List<object>();
            for (int j = 0; j < cols; j++)
            {
                rowArray.Add(array[i, j]);
            }

            dictionary[i + 1] = rowArray;
        }

        return dictionary;
    }

    //二维数组字典化-首列为Key，0开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstKey(object[,] array)
    {
        var dict = new Dictionary<string, List<string>>();
        for (int i = 0; i < array.GetLength(0); i++)
        {
            string key = array[i, 0]?.ToString();

            var row = new List<string>();
            for (int j = 0; j < array.GetLength(1); j++)
                row.Add(array[i, j]?.ToString());
            dict[key] = row;
        }
        return dict;
    }

    //二维数组字典化-首列为Key，0开始
    public static Dictionary<string, string> TwoDArrayToDicFirstKeyStr(object[,] array)
    {
        var dict = new Dictionary<string, string>();
        for (int i = 0; i < array.GetLength(0); i++)
        {
            string key = array[i, 0].ToString();

            string row = String.Empty;
            for (int j = 0; j < array.GetLength(1); j++)
                row = string.Join("#", array[i, j]?.ToString());
            dict[key] = row;
        }
        return dict;
    }

    //二维数组字典化-首行为Key，0开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstRowKey(object[,] array)
    {
        var dict = new Dictionary<string, List<string>>();
        for (int j = 0; j < array.GetLength(1); j++)
        {
            string key = array[0, j]?.ToString();

            var col = new List<string>();

            // 不保存表头数据
            for (int i = 1; i < array.GetLength(0); i++)
                col.Add(array[i, j]?.ToString());
            dict[key] = col;
        }
        return dict;
    }

    //二维数组字典化-首列为Key,ExcelRange对象，1开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstKey1(object[,] array)
    {
        var dict = new Dictionary<string, List<string>>();
        for (int i = 1; i <= array.GetLength(0); i++)
        {
            string key = array[i, 1]?.ToString();
            if (key == string.Empty)
                continue;
            var row = new List<string>();
            for (int j = 1; j <= array.GetLength(1); j++)
                row.Add(array[i, j]?.ToString());
            dict[key] = row;
        }
        return dict;
    }

    //二维数组字典化-首列为Key,ExcelRange对象，1开始
    public static Dictionary<string, string> TwoDArrayToDicFirstKeyStr1(object[,] array)
    {
        var dict = new Dictionary<string, string>();
        for (int i = 1; i <= array.GetLength(0); i++)
        {
            string key = array[i, 1].ToString();

            string row = String.Empty;
            for (int j = 1; j <= array.GetLength(1); j++)
                row = string.Join("#", array[i, j]?.ToString());
            dict[key] = row;
        }
        return dict;
    }

    //二维数组字典化-首行为Key,ExcelRange对象，1开始
    public static Dictionary<string, List<string>> TwoDArrayToDictionaryFirstRowKey1(
        object[,] array
    )
    {
        var dict = new Dictionary<string, List<string>>();
        for (int j = 1; j <= array.GetLength(1); j++)
        {
            string key = array[1, j].ToString();

            var col = new List<string>();
            for (int i = 1; i <= array.GetLength(0); i++)
                col.Add(array[i, j]?.ToString());
            dict[key] = col;
        }
        return dict;
    }

    //二维数组转二维字典
    public static Dictionary<(object, object), string> Array2DToDic2D(
        int rowCount,
        int colCount,
        object[,] modelRangeValue
    )
    {
        var modelValue = new Dictionary<(object, object), string>();
        for (int row = 2; row <= rowCount; row++)
        {
            for (int col = 2; col <= colCount; col++)
            {
                var rowIndex = modelRangeValue[row, 1];
                var colIndex = modelRangeValue[1, col];
                if (rowIndex == null || colIndex == null)
                {
                    MessageBox.Show(@"模版表中表头有空值，请检查模版数据是否正确！");
                    return null;
                }

                string value = modelRangeValue[row, col]?.ToString() ?? "";
                modelValue[(rowIndex, colIndex)] = value;
            }
        }

        return modelValue;
    }

    public static Dictionary<(object, object), string> Array2DToDic2D0(
        int rowCount,
        int colCount,
        object[,] modelRangeValue
    )
    {
        object[,] modelRangeValues = (object[,])modelRangeValue;
        var modelValue = new Dictionary<(object, object), string>();
        for (int row = 1; row < rowCount; row++)
        {
            for (int col = 1; col < colCount; col++)
            {
                var rowIndex = modelRangeValues[row, 0];
                var colIndex = modelRangeValues[0, col];
                if (rowIndex == null || colIndex == null)
                {
                    MessageBox.Show(@"模版表中表头有空值，请检查模版数据是否正确！");
                    return null;
                }

                string value = modelRangeValues[row, col]?.ToString() ?? "";
                modelValue[(rowIndex, colIndex)] = value;
            }
        }

        return modelValue;
    }

    //字典二维数组化
    public static object[,] DictionaryTo2DArray<TKey, TValue>(
        Dictionary<TKey, List<TValue>> dictionary,
        int? maxRows = null,
        int? maxCols = null
    )
    {
        int rows = maxRows ?? dictionary.Count;
        int cols = maxCols ?? (dictionary.Values.Max(list => list.Count) + 1);

        object[,] array2D = new object[rows, cols];

        int row = 0;
        foreach (var kvp in dictionary)
        {
            if (row >= rows)
                break;

            for (int col = 0; col < Math.Min(kvp.Value.Count, cols); col++)
            {
                array2D[row, col] = kvp.Value[col];
            }

            row++;
        }

        return array2D;
    }

    //字典二维数组化-带Key(数据模版专用）
    public static object[,] DictionaryTo2DArrayKey<TKey, TValue>(
        Dictionary<TKey, List<TValue>> dictionary,
        int maxRows,
        int maxCols
    )
    {
        object[,] array2D = new object[maxRows, maxCols];

        int row = 0;
        foreach (var kvp in dictionary)
        {
            bool isFirstValue = true;
            foreach (var value in kvp.Value)
            {
                array2D[row, 0] = value;
                array2D[row, 1] = null;
                array2D[row, 2] = isFirstValue ? kvp.Key : null;
                isFirstValue = false;
                row++;
            }
        }

        return array2D;
    }

    //二维数据字符串连接化缩短列数
    public static object[,] ConvertToCommaSeparatedArray(object[,] array2D)
    {
        int rows = array2D.GetLength(0);
        int cols = array2D.GetLength(1);

        string[,] newArray2D = new string[rows, 1];

        for (int i = 0; i < rows; i++)
        {
            List<string> rowElements = new List<string>();
            for (int j = 0; j < cols; j++)
            {
                rowElements.Add(array2D[i, j]?.ToString() ?? "null");
            }

            newArray2D[i, 0] = string.Join(",", rowElements);
        }

        return newArray2D;
    }

    //字典里随机选择若干条数据
    public static Dictionary<int, List<int>> RandChooseDataFormDictionary(
        Dictionary<int, List<int>> sourceDic,
        int chooseCount
    )
    {
        // 将字典的键转换为列表
        List<int> keys = sourceDic.Keys.ToList();
        //不够则有多少取多少
        chooseCount = Math.Min(chooseCount, keys.Count);
        // 使用随机数生成器随机选择 N个键
        Random random = new Random();
        List<int> selectedKeys = keys.OrderBy(x => random.Next()).Take(chooseCount).ToList();
        // 使用选中的键从字典中获取对应的值
        Dictionary<int, List<int>> selectedData = new Dictionary<int, List<int>>();
        foreach (int key in selectedKeys)
        {
            selectedData[key] = sourceDic[key];
        }

        return selectedData;
    }

    //二维数组去重
    public static object[,] CleanRepeatValue(
        object[,] array,
        int index,
        bool isRow,
        int baseIndex,
        bool emptyFilter = true
    )
    {
        var seen = new HashSet<object>(); // 用于存储已出现的基准值
        var tempResult = new List<object[]>(); // 临时存储去重后的结果

        int rows = array.GetLength(0); // 获取行数
        int cols = array.GetLength(1); // 获取列数

        // 检查 baseIndex 是否为 0 或 1
        if (baseIndex != 0 && baseIndex != 1)
        {
            throw new ArgumentException("Base index must be 0 or 1.", nameof(baseIndex));
        }

        // 检查 index 是否超出数组的范围 (根据 baseIndex 调整)
        if (
            index < baseIndex
            || (isRow && index >= cols + baseIndex)
            || (!isRow && index >= rows + baseIndex)
        )
        {
            throw new ArgumentOutOfRangeException(
                nameof(index),
                "Index is outside the bounds of the array."
            );
        }

        // 遍历方向控制
        int outerLoop = isRow ? cols : rows;
        int innerLoop = isRow ? rows : cols;

        for (int i = baseIndex; i < outerLoop + baseIndex; i++) // 根据 baseIndex 调整循环起点
        {
            // 如果 baseIndex 是 1，直接使用 index 和 i；如果是 0，减去 baseIndex
            var key = isRow
                ? array[
                    baseIndex == 1 ? index : index - baseIndex,
                    baseIndex == 1 ? i : i - baseIndex
                ]
                : array[
                    baseIndex == 1 ? i : i - baseIndex,
                    baseIndex == 1 ? index : index - baseIndex
                ];

            if (emptyFilter)
            {
                // 过滤掉 null 和空字符串
                if (key == null || (key is string str && string.IsNullOrWhiteSpace(str)))
                {
                    continue; // 跳过空值
                }
            }

            if (!seen.Contains(key))
            {
                seen.Add(key);
                var row = new object[innerLoop];

                for (int j = baseIndex; j < innerLoop + baseIndex; j++) // 根据 baseIndex 调整循环起点
                {
                    // 检查是否超出数组边界
                    if (
                        isRow && (j - baseIndex >= rows || i - baseIndex >= cols)
                        || !isRow && (i - baseIndex >= rows || j - baseIndex >= cols)
                    )
                    {
                        throw new IndexOutOfRangeException(
                            $"Index out of bounds: i={i}, j={j}, rows={rows}, cols={cols}"
                        );
                    }

                    // 如果按行去重，保留列的值；否则保留行的值
                    row[j - baseIndex] = isRow
                        ? array[
                            baseIndex == 1 ? j : j - baseIndex,
                            baseIndex == 1 ? i : i - baseIndex
                        ]
                        : array[
                            baseIndex == 1 ? i : i - baseIndex,
                            baseIndex == 1 ? j : j - baseIndex
                        ];
                }

                tempResult.Add(row);
            }
        }

        // 将临时结果转换为二维数组
        var result = new object[tempResult.Count, innerLoop];
        for (int i = 0; i < tempResult.Count; i++)
        {
            for (int j = 0; j < innerLoop; j++)
            {
                result[i, j] = tempResult[i][j];
            }
        }

        return result;
    }

    //二维数组复制到剪切板
    public static void CopyArrayToClipboard(object[,] array)
    {
        // 获取数组的行数和列数
        int rows = array.GetLength(0);
        int cols = array.GetLength(1);

        // 构建制表符分隔的字符串
        StringBuilder sb = new StringBuilder();

        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                if (array[i, j] != null)
                {
                    sb.Append(array[i, j].ToString());
                }

                // 如果不是最后一列，添加制表符
                if (j < cols - 1)
                {
                    sb.Append("\t");
                }
            }

            // 如果不是最后一行，添加换行符
            if (i < rows - 1)
            {
                sb.AppendLine();
            }
        }

        // 将字符串复制到剪贴板
        Clipboard.SetText(sb.ToString());
    }

    //数组变为二维化字符串
    public static string ArrayToArrayStr(object selectValue)
    {
        var resultStr = string.Empty;

        if (selectValue is object[,])
        {
            // 如果是二维数组
            var values = (object[,])selectValue;
            int rows = values.GetLength(0); // 获取行数
            int cols = values.GetLength(1); // 获取列数

            // 用 StringBuilder 拼接字符串
            var result = new System.Text.StringBuilder();

            for (int i = 1; i <= rows; i++) // 遍历每一行
            {
                for (int j = 1; j <= cols; j++) // 遍历每一列
                {
                    var cellValue = values[i, j] ?? ""; // 获取单元格值，处理空值
                    result.Append(cellValue.ToString()); // 拼接单元格值

                    if (j < cols)
                    {
                        result.Append(","); // 列之间用逗号分隔
                    }
                }

                if (i < rows)
                {
                    result.AppendLine(); // 行之间换行
                }
            }

            resultStr = result.ToString();
        }
        else
        {
            // 如果是单个值
            resultStr = selectValue.ToString();
        }

        return resultStr;
    }

    //Excel多选Range合并为二维数组

    public static object[,] MergeRanges(object[] areas, bool mergeByRow)
    {
        int totalRows = 0;
        int totalCols = 0;

        // 计算合并后的数组大小
        foreach (var area in areas)
        {
            object[,] areaValues = (object[,])area;
            if (mergeByRow)
            {
                totalRows += areaValues.GetLength(0); // 累加行数
                totalCols = Math.Max(totalCols, areaValues.GetLength(1)); // 取最大列数
            }
            else
            {
                totalCols += areaValues.GetLength(1); // 累加列数
                totalRows = Math.Max(totalRows, areaValues.GetLength(0)); // 取最大行数
            }
        }

        // 创建合并后的二维数组
        object[,] mergedArray = new object[totalRows, totalCols];

        // 按行或按列合并数据
        if (mergeByRow)
        {
            int currentRow = 0;
            foreach (var area in areas)
            {
                object[,] areaValues = (object[,])area;
                int areaRows = areaValues.GetLength(0);
                int areaCols = areaValues.GetLength(1);

                for (int i = 0; i < areaRows; i++)
                {
                    for (int j = 0; j < areaCols; j++)
                    {
                        mergedArray[currentRow + i, j] = areaValues[i + 1, j + 1];
                    }
                }

                currentRow += areaRows; // 更新当前行位置
            }
        }
        else
        {
            int currentCol = 0;
            foreach (var area in areas)
            {
                object[,] areaValues = (object[,])area;
                int areaRows = areaValues.GetLength(0);
                int areaCols = areaValues.GetLength(1);

                for (int i = 0; i < areaRows; i++)
                {
                    for (int j = 0; j < areaCols; j++)
                    {
                        mergedArray[i, currentCol + j] = areaValues[i + 1, j + 1];
                    }
                }

                currentCol += areaCols; // 更新当前列位置
            }
        }

        return mergedArray;
    }

    // 查找二维数组中的值，返回行和列的元组
    public static (int, int) FindValueIn2DArray(object[,] array, object value)
    {
        // 获取数组的行和列的起始索引
        int rowStart = array.GetLowerBound(0);
        int colStart = array.GetLowerBound(1);

        // 获取数组的行和列的结束索引
        int rowEnd = array.GetUpperBound(0);
        int colEnd = array.GetUpperBound(1);

        // 遍历数组
        for (int row = rowStart; row <= rowEnd; row++) // 遍历行
        {
            for (int col = colStart; col <= colEnd; col++) // 遍历列
            {
                // 检查是否为空值，并进行比较
                if (array[row, col] != null && array[row, col].ToString() == value.ToString())
                {
                    return (row, col); // 找到值，返回行和列
                }
            }
        }

        return (-1, -1); // 未找到值，返回 (-1, -1)
    }

    #region 自定义数组类型判断

    //检查并解析一维数组
    public static bool IsValidArray(string input, out object[] array)
    {
        array = null;
        if (input.StartsWith("[") && input.EndsWith("]"))
        {
            string content = input.Substring(1, input.Length - 2);
            array = content.Split(',').Select(s => (object)s.Trim()).ToArray();
            return true;
        }

        return false;
    }

    //检查并解析二维数组
    public static bool IsValidArray(string input, out object[][] array)
    {
        array = null;
        // 使用正则表达式验证二维数组的格式
        if (
            Regex.IsMatch(
                input,
                @"^\[\[(?:[^\[\]]+,\s*)*[^\[\]]+\](?:,\s*\[(?:[^\[\]]+,\s*)*[^\[\]]+\])*\]$"
            )
        )
        {
            // 去掉最外层的方括号
            input = input.Trim('[', ']');

            // 分割每一行
            var rows = input.Split(new[] { "],[" }, StringSplitOptions.None);

            // 去掉每一行的方括号
            rows = rows.Select(row => row.Trim('[', ']')).ToArray();

            // 转换为二维数组
            array = rows.Select(row =>
                    row.Split(',').Select(value => (object)value.Trim()).ToArray()
                )
                .ToArray();
            return true;
        }

        return false;
    }

    //检查一维数组中的元素是否为指定类型
    public static bool IsArrayOfType(object[] array, Type type)
    {
        if (array == null || type == null)
        {
            return false;
        }

        foreach (var element in array)
        {
            if (element == null)
            {
                return false;
            }

            try
            {
                // 尝试将元素转换为目标类型
                Convert.ChangeType(element, type);
            }
            catch (InvalidCastException)
            {
                return false;
            }
        }

        return true;
    }

    //检查二维数组中的元素是否为指定类型
    public static bool IsArrayOfType(object[][] array, Type type)
    {
        if (array == null || type == null)
        {
            return false;
        }

        foreach (var row in array)
        {
            foreach (var element in row)
            {
                if (element == null)
                {
                    return false;
                }

                try
                {
                    // 尝试将元素转换为目标类型
                    Convert.ChangeType(element, type);
                }
                catch (InvalidCastException)
                {
                    return false;
                }
            }
        }

        return true;
    }

    // 合并二维数组
    public static object[,] Merge2DArrays0(object[,] array1, object[,] array2)
    {
        // 合并数组
        int rowCount = array1.GetLength(0);
        int baseColCount = array1.GetLength(1);
        int tagColCount = array2.GetLength(1);

        if (rowCount != array2.GetLength(0))
        {
            throw new InvalidOperationException("两个数组的行数不一致，无法合并！");
        }

        object[,] mergedArray = new object[rowCount, baseColCount + tagColCount];

        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < baseColCount; col++)
            {
                mergedArray[row, col] = array1[row, col];
            }

            for (int col = 0; col < tagColCount; col++)
            {
                mergedArray[row, baseColCount + col] = array2[row, col];
            }
        }

        // 输出合并后的数组（示例）
        Debug.Print("合并成功！");
        return mergedArray;
    }

    public static object[,] Merge2DArrays1(object[,] array1, object[,] array2)
    {
        // 合并数组
        int rowCount = array1.GetLength(0);
        int baseColCount = array1.GetLength(1);
        int tagColCount = array2.GetLength(1);

        if (rowCount != array2.GetLength(0))
        {
            throw new InvalidOperationException("两个数组的行数不一致，无法合并！");
        }

        object[,] mergedArray = new object[rowCount, baseColCount + tagColCount];

        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < baseColCount; col++)
            {
                mergedArray[row, col] = array1[row + 1, col + 1];
            }

            for (int col = 0; col < tagColCount; col++)
            {
                mergedArray[row, baseColCount + col] = array2[row + 1, col + 1];
            }
        }

        // 输出合并后的数组（示例）
        Debug.Print("合并成功！");
        return mergedArray;
    }
    #endregion

    //查找资源文件
    public static Dictionary<string, List<string>> FindResourceFile(
        Dictionary<string, List<string>> longNumbers,
        string searchFolder
    )
    {
        // 线程安全字典：存储并行查找的结果
        var tempDict = new ConcurrentDictionary<string, List<string>>();

        var searchOptions = new EnumerationOptions
        {
            MatchCasing = MatchCasing.CaseInsensitive,
            RecurseSubdirectories = true
        };

        // **并行遍历字典的 Key-Value 对**
        Parallel.ForEach(
            longNumbers,
            kvp =>
            {
                string dictKey = kvp.Key; // 原始 Key
                List<string> values = kvp.Value; // 该 Key 关联的 List<string>

                if (values.Count < 3)
                    return; // 确保 values 至少有 3 个元素

                string value1 = values[0]; // 第 1 个值
                string subFolder = values[1]; // 用于拼接路径
                string searchNum = values[2]; // 作为图片名称查找

                string newSearchPath = Path.Combine(searchFolder, subFolder); // 拼接路径
                List<string> foundImages = new List<string>();

                if (Directory.Exists(newSearchPath)) // 确保目录存在
                {
                    var files = Directory.EnumerateFiles(
                        newSearchPath,
                        $"{searchNum}.png",
                        searchOptions
                    );
                    foundImages.AddRange(files);
                }

                // **存入 Key，包含 (value1, searchNum, 图片路径)**
                if (foundImages.Count > 0)
                {
                    tempDict[dictKey] = new List<string> { value1, searchNum };
                    tempDict[dictKey].AddRange(foundImages); // 添加所有找到的图片路径
                }
                else
                {
                    tempDict[dictKey] = new List<string> { value1, searchNum }; // 即使没找到，也存储基础数据
                }
            }
        );

        // **保证返回的 Dictionary 顺序与 longNumbers 一致**
        var orderedDict = new Dictionary<string, List<string>>();
        foreach (var key in longNumbers.Keys)
        {
            if (tempDict.TryGetValue(key, out var value))
            {
                orderedDict[key] = value;
            }
            else
            {
                orderedDict[key] = new List<string> { longNumbers[key][0], longNumbers[key][2] }; // 确保所有 Key 都存在
            }
        }

        return orderedDict;
    }

    // 检查Excel单元格值的合法性
    public static List<(string, int, int, string, string, string)> ExcelCellValueFormatCheck(
        string cellValue,
        string typeCell,
        string sheetName,
        string filePath,
        int rowIndex,
        int colIndex
    )
    {
        var config = new GlobalVariable();
        var normalCharactersCheck = config.NormaKeyList;
        var specialCharactersCheck = config.SpecialKeyList;
        var coupleCharactersCheck = config.CoupleKeyList;

        var sourceData = new List<(string, int, int, string, string, string)>();

        if (cellValue != null)
        {
            if (
                normalCharactersCheck.Any(c => cellValue.Contains(c))
                && !typeCell.Contains("string")
            )
            {
                sourceData.Add(
                    (cellValue, rowIndex + 1, colIndex + 1, sheetName, "多逗号或中文逗号", filePath)
                );
            }

            if (
                specialCharactersCheck.Any(c => cellValue.Contains(c))
                && !typeCell.Contains("string")
            )
            {
                sourceData.Add((cellValue, rowIndex + 1, colIndex + 1, sheetName, "少逗号", filePath));
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
                        (cellValue, rowIndex + 1, colIndex + 1, sheetName, "括号问题", filePath)
                    );
                    break;
                }

                if (leftString == "\"")
                {
                    int isDouble = leftStringCount % 2;
                    if (isDouble != 0)
                    {
                        sourceData.Add(
                            (cellValue, rowIndex + 1, colIndex + 1, sheetName, "双引号问题", filePath)
                        );
                        break;
                    }
                }
            }
        }

        return sourceData;
    }
}
