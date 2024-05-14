using OfficeOpenXml;
using System.Threading.Tasks;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Data.OleDb;
using ExcelReference = ExcelDna.Integration.ExcelReference;
using System.Text.RegularExpressions;

// ReSharper disable All
#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类
/// </summary>
public class PubMetToExcel
{
    #region EPPlus与Excel

    //EPPlus创建新Excel表格或获取已经存在的表格
    public static List<(string, string, string)> OpenOrCreatExcelByEpPlus(string excelFilePath, string excelName,
        out ExcelWorksheet sheet, out ExcelPackage excel)
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

        var errorList = SetExcelObjectEpPlus(excelFilePath, excelName + @".xlsx", out sheet, out excel);
        return errorList;
    }

    //EPPlus创建Excel对象
    public static List<(string, string, string)> SetExcelObjectEpPlus(dynamic excelPath, dynamic excelName,
        out ExcelWorksheet sheet, out ExcelPackage excel)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        sheet = null;
        excel = null;
        string errorExcelLog;
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

    public static List<int> MergeExcelCol(object[,] sourceRangeValue, ExcelWorksheet targetSheet,
        object[,] targetRangeTitle, object[,] sourceRangeTitle)
    {
        var targetColList = new List<int>();
        var defaultCol = targetSheet.Dimension.End.Column;
        var beforTargetCol = defaultCol;
        for (var c = 0; c < sourceRangeValue.GetLength(1); c++)
        {
            //target中找列
            var sourceCol = sourceRangeValue[1, c];
            if (sourceCol == null) sourceCol = "";

            //获取目标单元格填写数值的位置，默认位置未最后一行
            var targetCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, sourceCol.ToString());
            if (targetCol == -1)
            {
                targetSheet.InsertColumn(beforTargetCol + 1, 1);
                targetCol = beforTargetCol + 1;
            }

            beforTargetCol = targetCol;
            for (var i = 0; i < targetRangeTitle.GetLength(0); i++)
            {
                var targetTitle = targetRangeTitle[i, 0];
                if (targetTitle == null) targetTitle = "";

                for (var j = 0; j < sourceRangeTitle.GetLength(0); j++)
                {
                    var sourceTitle = sourceRangeTitle[j, 0];
                    if (sourceTitle == null) sourceTitle = "";

                    //target中找列
                    if (targetTitle.ToString() == sourceTitle.ToString())
                    {
                        var sourceValue = sourceRangeValue[c, j];
                        if (sourceValue == null) sourceValue = "";

                        var targetCell = targetSheet.Cells[targetCol, i + 1];
                        targetCell.Value = sourceValue;
                    }
                }
            }

            targetColList.Add(targetCol);
        }

        return targetColList;
    }

    public static List<int> MergeExcel(object[,] sourceRangeValue, ExcelWorksheet targetSheet,
        object[,] targetRangeTitle, object[,] sourceRangeTitle)
    {
        var targetRowList = new List<int>();
        var defaultRow = targetSheet.Dimension.End.Row;
        var beforTargetRow = defaultRow;
        for (var r = 0; r < sourceRangeValue.GetLength(0); r++)
        {
            //target中找行
            var sourceRow = sourceRangeValue[r, 1];
            if (sourceRow == null) sourceRow = "";

            //获取目标单元格填写数值的位置，默认位置未最后一行
            var targetRow = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, sourceRow.ToString());
            if (targetRow == -1)
            {
                targetSheet.InsertRow(beforTargetRow + 1, 1);
                targetRow = beforTargetRow + 1;
            }

            beforTargetRow = targetRow;
            for (var i = 0; i < targetRangeTitle.GetLength(1); i++)
            {
                var targetTitle = targetRangeTitle[0, i];
                if (targetTitle == null) targetTitle = "";

                for (var j = 0; j < sourceRangeTitle.GetLength(1); j++)
                {
                    var sourceTitle = sourceRangeTitle[0, j];
                    if (sourceTitle == null) sourceTitle = "";

                    //target中找列
                    if (targetTitle.ToString() == sourceTitle.ToString())
                    {
                        var sourceValue = sourceRangeValue[r, j];
                        if (sourceValue == null) sourceValue = "";

                        var targetCell = targetSheet.Cells[targetRow, i + 1];
                        targetCell.Value = sourceValue;
                    }
                }
            }

            targetRowList.Add(targetRow);
        }

        return targetRowList;
    }

    #endregion

    #region C-API与Excel

    [ExcelFunction(IsHidden = true)]
    //通过C-API的方式读取打开当前活动Excel表格各个Sheet的数据
    public static object[,] ReadExcelDataC(string sheetName, int rowFirst, int rowLast, int colFirst, int colLast)
    {
        var sheet = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);
        var range = new ExcelReference(rowFirst, rowLast, colFirst, colLast, sheet.SheetId);
        var rangeValue = range.GetValue();
        //兼容range和cell获取数据变为二维数据
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

    //通过C-API的方式写入打开当前活动Excel表格各个Sheet的数据
    public static void WriteExcelDataC(string sheetName, int rowFirst, int rowLast, int colFirst, int colLast,
        object[,] rangeValue)
    {
        var sheet = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);
        var range = new ExcelReference(rowFirst, rowLast, colFirst, colLast, sheet.SheetId);
        ExcelAsyncUtil.QueueAsMacro(() => { range.SetValue(rangeValue); });
    }

    public static Task<(ExcelReference currentRange, string sheetName, string sheetPath)> GetCurrentExcelObjectC()
    {
        //因为Excel的异步问题导致return值只捕捉到第一次，所以使用TCS确保等待异步完成，进而获得正确的return
        var tcs = new TaskCompletionSource<(ExcelReference currentRange, string sheetName, string sheetPath)>();
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                // 获取当前工作簿、工作表选中单元格（[工作簿]工作表）
                var currentRange = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
                // 获取当前工作簿、工作表名称（[工作簿]工作表）
                var sheetName = (string)XlCall.Excel(XlCall.xlfGetDocument, 1);
                // 获取当前工作簿路径（不包含工作簿名称）
                var sheetPath = (string)XlCall.Excel(XlCall.xlfGetDocument, 2);
                // 处理获取结果的逻辑
                var result = (currentRange, sheetName, sheetPath);
                tcs.SetResult(result);
            }
            catch (Exception ex)
            {
                // 处理异常
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

    //Excel数据输出为List
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToList(dynamic workSheet)
    {
        Range dataRange = workSheet.UsedRange;
        // 读取数据到一个二维数组中
        object[,] rangeValue = dataRange.Value;
        // 获取行数和列数
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        // 定义工作表数据数组和表头数组
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        // 读取数据和表头
        //单线程
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

            if (row > 1) sheetData.Add(rowList);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    //Excel数据输出为List，自定义数据起始行列
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToListBySelf(dynamic workSheet,
        int dataRow, int dataCol, int headerRow, int headerCol)
    {
        Range dataRange = workSheet.UsedRange;
        // 读取数据到一个二维数组中
        object[,] rangeValue = dataRange.Value;
        // 获取行数和列数
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        // 定义工作表数据数组和表头数组
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        // 读取数据
        //单线程
        for (var row = dataRow; row <= rows; row++)
        {
            var rowList = new List<object>();
            for (var column = dataCol; column <= columns; column++)
            {
                var value = rangeValue[row, column];
                if (row == headerRow) //这个可能是冗余判断，暂时未发现问题
                    sheetHeaderCol.Add(value);
                else
                    rowList.Add(value);
            }

            if (row > 1) sheetData.Add(rowList);
        }

        //读取表头
        for (var column = headerCol; column <= columns; column++)
        {
            var value = rangeValue[headerRow, column];
            sheetHeaderCol.Add(value);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    //Excel数据输出为List，自定义数据起始-结束行、列，根据当前选择的单元格
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToListBySelfToEnd(
        dynamic workSheet,
        int dataRow, int dataCol, int headRow)
    {
        Range selectRange = NumDesAddIn.App.Selection;
        Range usedRange = workSheet.UsedRange;
        int dataRowEnd;
        int dataColEnd;
        //确定行，不确定列
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

        //确定列，不确定行
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
        // 读取数据到一个二维数组中
        //object[,] rangeValue = dataRange.Value;
        Range headRangeStart = workSheet.Cells[headRow, dataCol];
        Range headRangeEnd = workSheet.Cells[headRow, dataColEnd];
        Range headRange = workSheet.Range[headRangeStart, headRangeEnd];
        // 读取数据到一个二维数组中
        //object[,] headRangeValue = headRange.Value;
        //// 定义工作表数据数组和表头数组
        //var sheetData = new List<List<object>>();
        //var sheetHeaderCol = new List<object>();
        //// 读取数据
        //for (var row = 1; row <= dataRowEnd - dataRow + 1; row++)
        //{
        //    var rowList = new List<object>();
        //    for (var column = 1; column <= dataColEnd - dataCol + 1; column++)
        //    {
        //        var value = rangeValue[row, column];
        //        rowList.Add(value);
        //    }
        //    sheetData.Add(rowList);
        //}
        ////读取表头
        //for (var column = 1; column <= dataColEnd - dataCol + 1; column++)
        //{
        //    var value = headRangeValue[headRow, column];
        //    sheetHeaderCol.Add(value);
        //}
        //var excelData = (sheetHeaderCol, sheetData);
        var excelData = RangeToListByVsto(dataRange, headRange, headRow);
        return excelData;
    }

    //Excel数据输出为DataTable，无表头，EPPlus
    public static DataTable ExcelDataToDataTable(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var file = new FileInfo(filePath);
        using (var package = new ExcelPackage(file))
        {
            var dataTable = new DataTable();
            // 默认导入第一个 Sheet 的数据
            var worksheet = package.Workbook.Worksheets["Sheet1"] ?? package.Workbook.Worksheets[0];
            dataTable.TableName = worksheet.Name;
            //创建列，可以添加值作为列名
            for (var col = 1; col <= worksheet.Dimension.End.Column; col++) dataTable.Columns.Add();

            // 读取数据行
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

    //Excel数据输出为DataTable，无表头，OLeDb，几乎是EPPlus的两倍速度
    public static DataTable ExcelDataToDataTableOleDb(string filePath)
    {
        // Excel 连接字符串，根据 Excel 版本和文件类型进行调整
        var connectionString =
            $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
        var sheetName = "Sheet1";
        using (var connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();
                var dataTable = new DataTable();

                // 获取所有可用的工作表名称
                var schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                // 检查 Sheet1 是否存在
                if (schemaTable != null)
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        // ReSharper disable once PossibleNullReferenceException
                        if (row != null && row["TABLE_NAME"].ToString().Equals("Sheet1"))
                        {
                            sheetName = "Sheet1";
                            break;
                        }

                        sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();
                    }

                // 读取 Excel 表格数据
                using (var command = new OleDbCommand($"SELECT * FROM [{sheetName}]", connection))
                {
                    using (var adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }

                if (sheetName != null) dataTable.TableName = sheetName;
                return dataTable;
            }
            catch (Exception ex)
            {
                // 处理异常
                Debug.Print("读取 Excel 表格数据出现异常：" + ex.Message);
                return null;
            }
        }
    }

    public static List<(string, string, int, int, string, string)> FindDataInDataTable(string fileFullName,
        dynamic dataTable, string findValue)
    {
        var findValueList = new List<(string, string, int, int, string, string)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        var sheetName = dataTable.TableName.ToString().Replace("$", "");
        foreach (DataRow row in dataTable.Rows)
        foreach (DataColumn column in dataTable.Columns)
            //模糊查询
            if (isAll)
            {
                if (row != null && row[column].ToString().Contains(findValue))
                    findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2,
                        row.Table.Columns.IndexOf(column) + 1, row[1].ToString(), row[2].ToString()));
            }
            //精确查询
            else
            {
                if (row[column].ToString() == findValue)
                    findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2,
                        row.Table.Columns.IndexOf(column) + 1, row[1].ToString(), row[2].ToString()));
            }

        return findValueList;
    }

    public static List<(string, string, int, int, string, string)> FindDataInDataTableKey(string fileFullName,
        dynamic dataTable, string findValue, int key)
    {
        var findValueList = new List<(string, string, int, int, string, string)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        var sheetName = dataTable.TableName.ToString().Replace("$", "");
        foreach (DataRow row in dataTable.Rows)
            //模糊查询
            if (isAll)
            {
                if (row != null && row[key - 1].ToString().Contains(findValue))
                    findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2, key, row[1].ToString(),
                        row[2].ToString()));
            }
            //精确查询
            else
            {
                if (row[key - 1].ToString() == findValue)
                    findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2, key, row[1].ToString(),
                        row[2].ToString()));
            }

        return findValueList;
    }

    public static string[] PathExcelFileCollect(List<string> pathList, string fileSuffixName, string[] ignoreFileNames)
    {
        var files = new string[] { };
        foreach (var path in pathList)
        {
            var file = Directory.EnumerateFiles(path, fileSuffixName)
                .Where(file => !ignoreFileNames.Any(ignore => Path.GetFileName(file).Contains(ignore)))
                .ToArray();
            files = files.Concat(file).ToArray();
        }

        return files;
    }

    public static Dictionary<string, List<Tuple<object[,]>>> ExcelDataToDictionary(dynamic data, dynamic dicKeyCol,
        dynamic dicValueCol, int valueRowCount, int valueColCount = 1)
    {
        var dic = new Dictionary<string, List<Tuple<object[,]>>>();

        for (var i = 0; i < data.Count; i++)
        {
            var value = data[i][dicKeyCol];

            if (value == null) continue;

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
                    // 记录错误日志
                    errorLog += $"Error {repeatValue}: {ex.Message}\n";
                }
        }

        return errorLog;
    }

    public static string RepeatValue2(ExcelWorksheet sheet, int row, int col, List<string> repeatValue)
    {
        var errorLog = "";
        // 获取指定列的单元格数据
        var sourceValues = sheet.Cells[row, col, sheet.Dimension.End.Row, col].Select(c => c.Value.ToString()).ToList();

        // 生成索引List
        var indexList = new List<int>();
        foreach (var repeat in repeatValue)
        {
            // 查找存在的值所在的行
            var rowIndex = sourceValues.FindIndex(c => c == repeat);
            if (rowIndex == -1) continue;
            rowIndex += row;
            indexList.Add(rowIndex);
        }

        indexList.Sort();
        if (indexList.Count != 0)
        {
            //合并List
            var outputList = new List<List<int>>();
            var start = indexList[0];

            for (var i = 1; i < indexList.Count; i++)
                if (indexList[i] != indexList[i - 1] + 1)
                {
                    outputList.Add(new List<int>() { start, indexList[i - 1] });
                    start = indexList[i];
                }

            outputList.Add(new List<int>() { start, indexList[indexList.Count - 1] });
            // 翻转输出列表
            outputList.Reverse();
            // 删除要删除的行
            foreach (var rowToDelete in outputList)
                try
                {
                    sheet.DeleteRow(rowToDelete[0], rowToDelete[1] - rowToDelete[0] + 1);
                }
                catch (Exception)
                {
                    //可能因为之前的填充格式问题导致异常
                    // 记录错误日志
                    errorLog += $"Error {sheet.Name}:#行号{rowToDelete}背景格式问题，更改背景色重试\n";
                }
        }

        return errorLog;
    }

    public static List<(string, string, int, int)> ErrorKeyFromExcelAll(string rootPath, string errorValue)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
        var mainPath = newPath + @"\Excels\Tables\";
        var files1 = Directory.EnumerateFiles(mainPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var langPath = newPath + @"\Excels\Localizations\";
        var files2 = Directory.EnumerateFiles(langPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var uiPath = newPath + @"\Excels\UIs\";
        var files3 = Directory.EnumerateFiles(uiPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var files = files1.Concat(files2).Concat(files3).ToArray();

        var targetList = new List<(string, string, int, int)>();
        var currentCount = 0;
        var count = files.Length;
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");
        foreach (var file in files)
        {
            try
            {
                // 使用 EPPlus 打开 Excel 文件进行操作
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    var wk = package.Workbook;
                    try
                    {
                        var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
                        for (var row = 4; row <= sheet.Dimension.End.Row; row++)
                        {
                            // 获取当前行的单元格数据
                            var cellValue = sheet.Cells[row, col].Value;
                            if (!isAll)
                            {
                                // 全词
                                if (cellValue != null && cellValue.ToString() == errorValue)
                                {
                                    // 返回该单元格的行地址
                                    var cellAddress = new ExcelCellAddress(row, col);
                                    var cellCol = cellAddress.Column;
                                    var cellRow = cellAddress.Row;
                                    targetList.Add((file, sheet.Name, cellRow, cellCol));
                                }
                            }
                            else
                            {
                                // 模糊
                                if (cellValue != null && cellValue.ToString().Contains(errorValue))
                                {
                                    // 返回该单元格的行地址
                                    var cellAddress = new ExcelCellAddress(row, col);
                                    var cellCol = cellAddress.Column;
                                    var cellRow = cellAddress.Row;
                                    targetList.Add((file, sheet.Name, cellRow, cellCol));
                                }
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch
            {
                continue;
            }

            currentCount++;
            //wk.Properties.Company = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
            NumDesAddIn.App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
        }

        return targetList;
    }

    public static List<(string, string, int, int)> ErrorKeyFromExcelAllMultiThread(string rootPath, string errorValue)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
        var mainPath = newPath + @"\Excels\Tables\";
        var files1 = Directory.EnumerateFiles(mainPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var langPath = newPath + @"\Excels\Localizations\";
        var files2 = Directory.EnumerateFiles(langPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var uiPath = newPath + @"\Excels\UIs\";
        var files3 = Directory.EnumerateFiles(uiPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var files = files1.Concat(files2).Concat(files3).ToArray();

        var targetList = new List<(string, string, int, int)>();
        var options = new ParallelOptions
            { MaxDegreeOfParallelism = Environment.ProcessorCount }; // 设置并行处理的最大线程数为处理器核心数

        Parallel.ForEach(files, options, file =>
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    try
                    {
                        var wk = package.Workbook;
                        var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
                        for (var row = 4; row <= sheet.Dimension.End.Row; row++)
                        {
                            var cellValue = sheet.Cells[row, col].Value;
                            if (cellValue != null && cellValue != null && cellValue.ToString().Contains(errorValue))
                            {
                                var cellAddress = new ExcelCellAddress(row, col);
                                var cellCol = cellAddress.Column;
                                var cellRow = cellAddress.Row;
                                targetList.Add((file, sheet.Name, cellRow, cellCol));
                            }
                        }
                    }
                    catch
                    {
                        // 处理异常
                    }
                }
            }
            catch
            {
                // ignored
            }
        });

        return targetList;
    }

    public static (string file, string Name, int cellRow, int cellCol) ErrorKeyFromExcelId(string rootPath,
        string errorValue)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
        var mainPath = newPath + @"\Excels\Tables\";
        var files1 = Directory.EnumerateFiles(mainPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var langPath = newPath + @"\Excels\Localizations\";
        var files2 = Directory.EnumerateFiles(langPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var uiPath = newPath + @"\Excels\UIs\";
        var files3 = Directory.EnumerateFiles(uiPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var files = files1.Concat(files2).Concat(files3).ToArray();

        var currentCount = 0;
        var count = files.Length;
        foreach (var file in files)
        {
            //过滤非配置表
            var fileName = Path.GetFileName(file);
            if (fileName.Contains("#")) continue;
            // 使用 EPPlus 打开 Excel 文件进行操作
            using (var package = new ExcelPackage(new FileInfo(file)))
            {
                try
                {
                    var wk = package.Workbook;
                    var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                    for (var col = 2; col <= 2; col++)
                    for (var row = 4; row <= sheet.Dimension.End.Row; row++)
                    {
                        // 获取当前行的单元格数据
                        var cellValue = sheet.Cells[row, col].Value;

                        // 如果找到了匹配的值
                        if (cellValue != null && cellValue.ToString() == errorValue)
                        {
                            // 返回该单元格的行地址
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
            //wk.Properties.Company = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
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

        if (a > 0) return ChangeExcelColChar(a - 1) + (char)(b + 65);

        return ((char)(b + 65)).ToString();
    }

    public static List<string> ReadWriteTxt(string filePath)
    {
        var textLineList = new List<string>();
        if (!File.Exists(filePath))
        {
            // 创建新的文本文件
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
            // 读取已存在的文本文件
            using var reader = new StreamReader(filePath);
            while (reader.ReadLine() is { } line) textLineList.Add(line);
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
            if (errorCell == "-1") continue;
            errorLog = errorLog + "【" + errorCell + "】" + errorExcelName + "#" + errorExcelLog + "\r\n";
        }

        return errorLog;
    }

    [ExcelFunction(IsHidden = true)]
    //数字转Excel字母列
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

    //打开指定Excel文件，并定位到指定sheet的指定单元格（Com）
    public static void OpenExcelAndSelectCell(string filePath, string sheetName, string cellAddress)
    {
        try
        {
            //验证文件名
            if (!File.Exists(filePath))
            {
                System.Windows.Forms.MessageBox.Show("文件不存在，请检查！");
                return;
            }

            // 打开指定路径的 Excel 文件
            var workbook = NumDesAddIn.App.Workbooks.Open(filePath);
            // 获取指定名称的工作表
            var worksheet = workbook.Sheets[sheetName];
            // 选择指定的单元格,非法则选择默认值A1
            var regex = new Regex(@"^[A-Za-z]+\d+$");
            var cellAddressDefault = "A1";
            if (cellAddress != null)
                if (regex.IsMatch(cellAddress))
                    cellAddressDefault = cellAddress;
            var cellRange = worksheet.Range[cellAddressDefault];
            worksheet.Select();
            cellRange.Select();
        }
        catch (Exception)
        {
            //异常处理
        }

        // 垃圾回收
        GC.Collect();
    }

    //List转换数据为Range数据（已开启的表格）（Com）
    public static void ListToArrayToRange(List<List<object>> targetList, dynamic workSheet, int startRow, int startCol)
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
        var targetRange = workSheet.Range[workSheet.Cells[startRow, startCol],
            workSheet.Cells[startRow + rowCount - 1, startCol + columnCount - 1]];
        targetRange.Value = targetDataArr;
    }

    //Range数据转List（Com）
    public static List<List<object>> RangeDataToList(object[,] rangeValue)
    {
        // 获取行数和列数
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);
        // 定义工作表数据数组和表头数组
        var sheetData = new List<List<object>>();
        // 读取数据
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

    //随机不重复值列表
    public static List<int> GenerateUniqueRandomList(int minValue, int maxValue, int baseValue)
    {
        var list = new List<int>();

        // 初始化列表
        for (var i = minValue; i <= maxValue; i++) list.Add(i + baseValue);

        // 使用 Fisher-Yates 洗牌算法生成随机不重复列表
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

    //List转数组
    public static object[,] ConvertListToArray(List<List<object>> listOfLists)
    {
        var rowCount = listOfLists.Count;
        var colCount = listOfLists.Count > 0 ? listOfLists[0].Count : 0;

        var twoDArray = new object[rowCount, colCount];

        for (var i = 0; i < rowCount; i++)
        {
            var innerList = listOfLists[i];

            for (var j = 0; j < colCount; j++) twoDArray[i, j] = innerList[j];
        }

        return twoDArray;
    }

    //VSTO内置在Range内查找特定值(第一个)的方法
    public static (int row, int column) FindValueInRangeByVsto(Range searchRange, object valueToFind)
    {
        // 使用 Find 方法在指定范围内查找特定值
        Range foundRange = searchRange.Find(valueToFind);
        // 如果找到了特定值
        if (foundRange != null)
        {
            // 返回找到的单元格的行号和列号
            return (foundRange.Row, foundRange.Column);
        }
        else
        {
            // 如果没有找到特定值，返回 (-1, -1)
            return (-1, -1);
        }
    }

    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) RangeToListByVsto(Range rangeData,
        Range rangeHeader, int headRow)
    {
        // 读取数据到一个二维数组中
        object[,] rangeValue = rangeData.Value;
        // 读取数据到一个二维数组中
        object[,] headRangeValue = rangeHeader.Value;
        // 定义工作表数据数组和表头数组
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        // 读取数据
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

        //读取表头
        for (var column = 1; column <= rangeValue.GetLength(1); column++)
        {
            var value = headRangeValue[headRow, column];
            sheetHeaderCol.Add(value);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    public static void TestEpPlus()
    {
        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //using (ExcelPackage package = new ExcelPackage(@"C:\Users\cent\Desktop\text.xlsx"))
        //{
        //    // Get the first worksheet
        //    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet2"];

        //    // Read the data from the worksheet
        //    var readValue = worksheet.Cells[1, 1].Value;

        //    for (int i = 1; i < 100001; i++)
        //    {
        //        var writeValue = readValue + "+New Value" + i;
        //        worksheet.Cells[i, 2].Value = writeValue;
        //    }
        //    package.Save();
        //}

        //string path = @"C:\Users\cent\Desktop\text.xlsx";
        //using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
        //{
        //    var wk = new XSSFWorkbook(stream);
        //    var sheet = wk.GetSheetAt(1);
        //    var readValue = sheet.GetRow(0).GetCell(0).ToString();
        //    for (int i = 0; i < 100001; i++)
        //    {
        //        var row = sheet.GetRow(i) ?? sheet.CreateRow(i);
        //        var cell = row.GetCell(1) ?? row.CreateCell(1);
        //        cell.SetCellValue(readValue + "+New Value" + i);
        //    }
        //    // 保存更改
        //    using (FileStream stream2 = new FileStream(path, FileMode.Create, FileAccess.Write))
        //    {
        //        wk.Write(stream2);
        //    }
        //}
    }
}
//// 自定义比较器来比较元组
//public class TupleEqualityComparer : IEqualityComparer<(string, string, string)>
//{
//    public bool Equals((string, string, string) x, (string, string, string) y)
//    {
//        return x.Item1 == y.Item1 && x.Item2 == y.Item2 && x.Item3 == y.Item3;
//    }

//    public int GetHashCode((string, string, string) obj)
//    {
//        return obj.Item1.GetHashCode() ^ obj.Item2.GetHashCode() ^ obj.Item3.GetHashCode();
//    }
//}