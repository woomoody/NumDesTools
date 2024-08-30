using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;
using ExcelReference = ExcelDna.Integration.ExcelReference;

// ReSharper disable All

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类
/// </summary>
public class PubMetToExcel
{
    #region EPPlus与Excel

    public static List<(string, string, string)> OpenOrCreatExcelByEpPlus(
        string excelFilePath,
        string excelName,
        out ExcelWorksheet sheet,
        out ExcelPackage excel
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        sheet = null;
        excel = null;
        string errorExcelLog;
        var errorList = new List<(string, string, string)>();
        string path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        string sheetRealName = "Sheet1";
        string excelRealName = excelName;
        if (excelName.Contains("#"))
        {
            var excelRealNameGroup = excelName.Split("#");
            excelRealName = excelRealNameGroup[0];
            sheetRealName = excelRealNameGroup[1];
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
            return errorList;
        }

        sheet ??= workBook.Worksheets[0];
        return errorList;
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
            var cellValue = sheet.Cells[row, col].Value;

            if (cellValue != null && cellValue.ToString() == searchValue)
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
        for (var row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            var cellValue = sheet.Cells[row, col].Value;

            if (cellValue != null && cellValue.ToString() == searchValue)
            {
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Row;
                return rowAddress;
            }
        }

        return -1;
    }

    #endregion

    #region C-API与Excel

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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

            if (value == null)
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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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

    [ExcelFunction(IsHidden = true)]
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

    [ExcelFunction(IsHidden = true)]
    public static string ErrorLogAnalysis(dynamic errorList, dynamic sheet)
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

    [ExcelFunction(IsHidden = true)]
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
                MessageBox.Show("文件不存在，请检查！");
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

    public static object[,] ConvertListToArray(List<List<object>> listOfLists)
    {
        var rowCount = listOfLists.Count;
        var colCount = listOfLists.Count > 0 ? listOfLists[0].Count : 0;

        var twoDArray = new object[rowCount, colCount];

        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            for (var j = 0; j < colCount; j++)
                twoDArray[i, j] = innerList[j];
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
    //二维数组转二维字典
    public static  Dictionary<(object, object), string>  Array2DToDic2D(int rowCount, int colCount, dynamic modelRangeValue)
    {
        var modelValue = new Dictionary<(object, object), string>();
        for (int row = 2; row <= rowCount; row++)
        {
            for (int col = 2; col <= colCount; col++)
            {
                string rowIndex = modelRangeValue[row, 1];
                string colIndex = modelRangeValue[1, col];
                if (rowIndex == null || colIndex == null)
                {
                    MessageBox.Show("模版表中表头有空值，请检查模版数据是否正确！");
                    return null ;
                }
                string value = modelRangeValue[row, col]?.ToString() ?? "";
                modelValue[(rowIndex, colIndex)] = value;
            }
        }
        return modelValue;
    }
    //字典二维数组化
    public static object[,] DictionaryTo2DArray(
        Dictionary<int, List<int>> dictionary,
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
    #region 自定义数组类型判断
    // 检查并解析一维数组
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

    // 检查并解析二维数组
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
            array = rows.Select(row => row.Split(',')
                    .Select(value => (object)value.Trim())
                    .ToArray())
                .ToArray();
            return true;
        }
        return false;
    }
    // 检查一维数组中的元素是否为指定类型
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

    // 检查二维数组中的元素是否为指定类型
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
    #endregion
    public static void TestEpPlus() { }
}
