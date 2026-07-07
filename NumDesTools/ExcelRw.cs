using System.Collections.Concurrent;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MiniExcelLibs;
using NumDesTools.Config;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;
using ExcelReference = ExcelDna.Integration.ExcelReference;

// ReSharper disable All

#pragma warning disable CA1416

namespace NumDesTools;

public static partial class PubMetToExcel
{
    // 延迟访问：禁止类加载期访问 COM，避免 Excel 未就绪时 NRE/COMException
    private static Workbook Wk => AppServices.App.ActiveWorkbook;

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
        if (excelName.Contains("Localizations"))
        {
            path = newPath + $@"\Excels\Localizations\{excelName}";
        }
        else
        {
            switch (excelName)
            {
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
        var fullPath = excelPath + @"\" + excelName;
        if (!File.Exists(fullPath))
        {
            PluginLog.Write($"[ExcelRw] 文件不存在: {fullPath}");
            return;
        }
        excel = new ExcelPackage(new FileInfo(fullPath));
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

        if (lastMatch != null)
        {
            return lastMatch.Start.Row;
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
        if (sheet is null)
        {
            PluginLog.Write(
                $"[GetExcelListObjects] sheet '{sheetName}' not found in {excelName} (path={excelPath})"
            );
            excel?.Dispose();
            return listObjectDataDic;
        }
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
        excel?.Dispose();
        return listObjectDataDic;
    }

    //获取指定表的名称表，当前
    public static ListObject GetExcelListObjects(string sheetName, string listName)
    {
        LogDisplay.RecordLine($"[{DateTime.Now}] 获取Excel ListObject: {sheetName} - {listName}");
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
                $"[{DateTime.Now}] 获取Excel ListObject: {sheetName} - {listName} 不存在-{e}"
            );
            throw;
        }
    }

    public static ListObject GetExcelListObjects2(Workbook workbook, string listName)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                foreach (ListObject listObj in sheet.ListObjects)
                {
                    if (listObj.Name == listName)
                        return listObj;
                }
            }
            catch (COMException) { }
        }
        return null;
    }

    public static ListObject GetExcelListObjectsBloor(Worksheet sheet, string listName)
    {
        try
        {
            foreach (ListObject listObj in sheet.ListObjects)
            {
                if (listObj.Name.Contains(listName))
                {
                    return listObj;
                }
            }
        }
        catch (COMException) { }

        return null;
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
        var workArea = AppServices.App.ActiveWindow;
        var targetXPoint = targetX * 1.67;
        var targetXPixels = workArea.PointsToScreenPixelsX((int)targetXPoint);
        return targetXPixels;
    }

    [ExcelFunction(IsHidden = true)]
    public static int ExcelRangePixelsY(double targetY)
    {
        var workArea = AppServices.App.ActiveWindow;
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
        Range selectRange = AppServices.App.Selection;
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
            if (worksheet?.Dimension is null)
                return dataTable;
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
                PluginLog.Write("读取 Excel 表格数据出现异常：" + ex.Message);
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

            if (string.IsNullOrEmpty(value?.ToString()))
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
        var errorLog = new StringBuilder();
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
                    errorLog.Append($"Error {repeatValue}: {ex.Message}\n");
                }
        }

        return errorLog.ToString();
    }

    public static string RepeatValue2(
        ExcelWorksheet sheet,
        int row,
        int col,
        List<string> repeatValue
    )
    {
        var errorLog = new StringBuilder();
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
                catch (Exception e)
                {
                    errorLog.Append(
                        $"Error {sheet.Name}:#行号{rowToDelete}背景格式问题，更改背景色重试 ({e.Message})\n"
                    );
                }
        }

        return errorLog.ToString();
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
            AppServices.App.StatusBar =
                "正在检查第" + currentCount + "/" + count + "个文件:" + file;
        }

        var tupleError = ("", "", 0, 0);
        return tupleError;
    }
}
