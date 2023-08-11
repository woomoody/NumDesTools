using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;
using System.Threading.Tasks;
using System.Data;
using DataTable = System.Data.DataTable;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.OleDb;
using NLua;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NumDesTools;
/// <summary>
/// 公共的Excel功能类
/// </summary>
public class PubMetToExcel
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
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
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToListBySelf(dynamic workSheet,int dataRow,int dataCol,int headerRow,int headerCol)
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
                if (row == headerRow)
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
    //Excel数据输出为DataTable，无表头，EPPlus
    public static DataTable ExcelDataToDataTable(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        FileInfo file = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(file))
        {
            var dataTable = new DataTable();
            // 默认导入第一个 Sheet 的数据
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"] ?? package.Workbook.Worksheets[0];
            dataTable.TableName = worksheet.Name;
            //创建列，可以添加值作为列名
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                dataTable.Columns.Add();
            }
            // 读取数据行
            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString();
                }
                dataTable.Rows.Add(dataRow);
            }
            return dataTable;
        }
    }
    //Excel数据输出为DataTable，无表头，OLeDb，几乎是EPPlus的两倍速度
    public static DataTable ExcelDataToDataTableOleDb(string filePath)
    {
        // Excel 连接字符串，根据 Excel 版本和文件类型进行调整
        string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES;'";
        var sheetName = "Sheet1";
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();
                DataTable dataTable = new DataTable();

                // 获取所有可用的工作表名称
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                // 检查 Sheet1 是否存在
                foreach (DataRow row in schemaTable.Rows)
                {
                    if (row["TABLE_NAME"].ToString().Equals("Sheet1"))
                    {
                        sheetName = "Sheet1";
                        break;
                    }
                    sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();
                }
                // 读取 Excel 表格数据
                using (OleDbCommand command = new OleDbCommand($"SELECT * FROM [{sheetName}]", connection))
                {
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(dataTable);
                    }
                }
                dataTable.TableName = sheetName;
                return dataTable;
            }
            catch (Exception ex)
            {
                // 处理异常
                Console.WriteLine("读取 Excel 表格数据出现异常：" + ex.Message);
                return null;
            }
        }
    }

    public static List<(string,string,int, int,string,string)> FindDataInDataTable(string fileFullName,dynamic dataTable, string findValue)
    {
        var findValueList = new List<(string,string,int, int,string,string)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        var sheetName = dataTable.TableName.ToString().Replace("$","");
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (DataColumn column in dataTable.Columns)
            {
                //模糊查询
                if (isAll)
                {
                    if (row[column].ToString().Contains(findValue))
                    {
                        findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2, row.Table.Columns.IndexOf(column) + 1, row[1].ToString(),row[2].ToString()));
                    }
                }
                //精确查询
                else
                {
                    if (row[column].ToString() == findValue)
                    {
                        findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2, row.Table.Columns.IndexOf(column) + 1, row[1].ToString(),row[2].ToString()));
                    }
                }
            }
        }
        return findValueList;
    }

    public static List<(string, string, int, int, string, string)> FindDataInDataTableKey(string fileFullName, dynamic dataTable, string findValue,int key)
    {
        var findValueList = new List<(string, string, int, int, string, string)>();
        var isAll = findValue.Contains("*");
        findValue = findValue.Replace("*", "");
        var sheetName = dataTable.TableName.ToString().Replace("$", "");
        foreach (DataRow row in dataTable.Rows)
        {
            //模糊查询
            if (isAll)
            {
                if (row[key-1].ToString().Contains(findValue))
                {
                    findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2, key, row[1].ToString(), row[2].ToString()));
                }
            }
            //精确查询
            else
            {
                if (row[key-1].ToString() == findValue)
                {
                    findValueList.Add((fileFullName, sheetName, row.Table.Rows.IndexOf(row) + 2, key, row[1].ToString(), row[2].ToString()));
                }
            }
        }
        return findValueList;
    }

    public static string[] PathExcelFileCollect(List<string> pathList, string fileSuffixName,string ignoreFileName)
    {
        string[] files = new string[] { };
        foreach (var path in pathList)
        {
            var file = Directory.EnumerateFiles(path, fileSuffixName)
                .Where(file => !Path.GetFileName(file).Contains(ignoreFileName))
                .ToArray();
            files =files.Concat(file).ToArray();
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

    public static string RepeatValue(ExcelWorksheet sheet,int row, int col,string repeatValue)
    {
        string errorLog ="";
        for (int r = sheet.Dimension.End.Row; r >= row; r--)
        {
            var colA = sheet.Cells[r, col].Value?.ToString();
            if (colA == repeatValue)
            {
                try
                {
                    sheet.DeleteRow(r);
                }
                catch(Exception ex)
                {
                    // 记录错误日志
                    errorLog += $"Error {repeatValue}: {ex.Message}\n";
                }
            }
        }
        return errorLog;
    }

    public static string RepeatValue2(ExcelWorksheet sheet, int row, int col, List<string> repeatValue)
    {
        string errorLog = "";
        // 获取指定列的单元格数据
        var sourceValues = sheet.Cells[row, col, sheet.Dimension.End.Row, col].Select(c => c.Value.ToString()).ToList(); 

        // 生成索引List
        var indexList = new List<int>();
        foreach (var repeat in repeatValue)
        {
            // 查找存在的值所在的行
            int rowIndex = sourceValues.FindIndex(c => c == repeat) ;
            if (rowIndex == -1) continue;
            rowIndex += row;
            indexList.Add(rowIndex);
        }
        indexList.Sort();
        if (indexList.Count != 0)
        {
            //合并List
            List<List<int>> outputList = new List<List<int>>();
            int start = indexList[0];

            for (int i = 1; i < indexList.Count; i++)
            {
                if (indexList[i] != indexList[i - 1] + 1)
                {
                    outputList.Add(new List<int>() { start, indexList[i - 1] });
                    start = indexList[i];
                }
            }
            outputList.Add(new List<int>() { start, indexList[indexList.Count - 1] });
            // 翻转输出列表
            outputList.Reverse();
            // 删除要删除的行
            foreach (var rowToDelete in outputList)
            {
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
        }
        return errorLog;
    }

    public static List<(string,string,int,int)>  ErrorKeyFromExcelAll(string rootPath, string errorValue)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
        var mainPath = newPath + @"\Excels\Tables\";
        string[] files1 = Directory.EnumerateFiles(mainPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var langPath = newPath + @"\Excels\Localizations\";
        string[] files2 = Directory.EnumerateFiles(langPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var uiPath = newPath + @"\Excels\UIs\";
        string[] files3 = Directory.EnumerateFiles(uiPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var files = files1.Concat(files2).Concat(files3).ToArray();

        var targetList = new List<(string, string, int, int)>();
        int currentCount = 0;
        var count = files.Length;
        var isAll = errorValue.Contains("*");
        errorValue = errorValue.Replace("*", "");
        foreach (var file in files)
        {
            try
            {
                // 使用 EPPlus 打开 Excel 文件进行操作
                using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                {
                    var wk = package.Workbook;
                    try
                    {
                        var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
                        {
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
            App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
        }
        return targetList;
    }

    public static List<(string, string, int, int)> ErrorKeyFromExcelAllMultiThread(string rootPath, string errorValue)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
        var mainPath = newPath + @"\Excels\Tables\";
        string[] files1 = Directory.EnumerateFiles(mainPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var langPath = newPath + @"\Excels\Localizations\";
        string[] files2 = Directory.EnumerateFiles(langPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var uiPath = newPath + @"\Excels\UIs\";
        string[] files3 = Directory.EnumerateFiles(uiPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var files = files1.Concat(files2).Concat(files3).ToArray();

        var targetList = new List<(string, string, int, int)>();
        var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }; // 设置并行处理的最大线程数为处理器核心数

        Parallel.ForEach(files, options, file =>
        {
            try
            {

                using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                {
                    try
                    {
                        var wk = package.Workbook;
                        var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
                        {
                            for (var row = 4; row <= sheet.Dimension.End.Row; row++)
                            {
                                var cellValue = sheet.Cells[row, col].Value;
                                if (cellValue != null && cellValue.ToString().Contains(errorValue))
                                {
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

    public static (string file, string Name, int cellRow, int cellCol) ErrorKeyFromExcelId(string rootPath, string errorValue)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
        var mainPath = newPath + @"\Excels\Tables\";
        string[] files1 = Directory.EnumerateFiles(mainPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var langPath = newPath + @"\Excels\Localizations\";
        string[] files2 = Directory.EnumerateFiles(langPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var uiPath = newPath + @"\Excels\UIs\";
        string[] files3 = Directory.EnumerateFiles(uiPath, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"))
            .ToArray();
        var files = files1.Concat(files2).Concat(files3).ToArray();

        int currentCount = 0;
        var count = files.Length;
        foreach (var file in files)
        {
            //过滤非配置表
            var fileName = Path.GetFileName(file);
            if (fileName.Contains("#")) continue;
            // 使用 EPPlus 打开 Excel 文件进行操作
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                var wk = package.Workbook;
                var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                for (var col = 2; col <= 2; col++)
                {
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
            }
            currentCount++;
            //wk.Properties.Company = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
            App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
        }
        var tupleError = ("", "", 0, 0);
        return tupleError;
    }

    public static Color GetCellBackgroundColor(Range cell)
    {
        Color color = Color.Empty;

        if (cell.Interior.Color != null)
        {
            object excelColor = cell.Interior.Color;
            if (excelColor is double)
            {
                double colorValue = (double)excelColor;
                int intValue = (int)colorValue;
                int red = intValue & 0xFF;
                int green = (intValue & 0xFF00) >> 8;
                int blue = (intValue & 0xFF0000) >> 16;
                color = Color.FromArgb(red, green, blue);
            }
        }
        return color;
    }

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
            using (StreamWriter writer = File.CreateText(filePath))
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
            while (reader.ReadLine() is { } line)
            {
                textLineList.Add(line);
            }
        }
        return textLineList;
    }

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

    public static List<int> MergeExcel(object[,] sourceRangeValue, ExcelWorksheet targetSheet, object[,] targetRangeTitle, object[,] sourceRangeTitle)
    {
        var targetRowList =new List<int>();
        int defaultRow = targetSheet.Dimension.End.Row;
        int beforTargetRow= defaultRow;
        for (int r = 0; r < sourceRangeValue.GetLength(0); r++)
        {
            //target中找行
            var sourceRow = sourceRangeValue[r, 1];
            if (sourceRow == null)
            {
                sourceRow = "";
            }
            //获取目标单元格填写数值的位置，默认位置未最后一行
            var targetRow = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, sourceRow.ToString());
            if (targetRow == -1)
            {
                targetSheet.InsertRow(beforTargetRow + 1,1);
                targetRow = beforTargetRow + 1;
            }
            beforTargetRow = targetRow;
            for (int i = 0; i < targetRangeTitle.GetLength(1); i++)
            {
                var targetTitle = targetRangeTitle[0, i];
                if (targetTitle == null)
                {
                    targetTitle = "";
                }
                for (int j = 0; j < sourceRangeTitle.GetLength(1); j++)
                {
                    var sourceTitle = sourceRangeTitle[0, j];
                    if (sourceTitle == null)
                    {
                        sourceTitle = "";
                    }
                    //target中找列
                    if (targetTitle.ToString() == sourceTitle.ToString())
                    {
                        var sourceValue = sourceRangeValue[r, j];
                        if (sourceValue == null)
                        {
                            sourceValue = "";
                        }
                        var targetCell = targetSheet.Cells[targetRow, i + 1];
                        targetCell.Value = sourceValue;
                    }
                }
            }
            targetRowList.Add(targetRow);
        }
        return targetRowList;
    }

    public static List<int> MergeExcelCol(object[,] sourceRangeValue, ExcelWorksheet targetSheet, object[,] targetRangeTitle, object[,] sourceRangeTitle)
    {
        var targetColList = new List<int>();
        int defaultCol = targetSheet.Dimension.End.Column;
        int beforTargetCol = defaultCol;
        for (int c = 0; c < sourceRangeValue.GetLength(1); c++)
        {
            //target中找列
            var sourceCol = sourceRangeValue[1, c];
            if (sourceCol == null)
            {
                sourceCol = "";
            }
            //获取目标单元格填写数值的位置，默认位置未最后一行
            var targetCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, sourceCol.ToString());
            if (targetCol == -1)
            {
                targetSheet.InsertColumn(beforTargetCol + 1, 1);
                targetCol = beforTargetCol + 1;
            }
            beforTargetCol = targetCol;
            for (int i = 0; i < targetRangeTitle.GetLength(0); i++)
            {
                var targetTitle = targetRangeTitle[i, 0];
                if (targetTitle == null)
                {
                    targetTitle = "";
                }
                for (int j = 0; j < sourceRangeTitle.GetLength(0); j++)
                {
                    var sourceTitle = sourceRangeTitle[j, 0];
                    if (sourceTitle == null)
                    {
                        sourceTitle = "";
                    }
                    //target中找列
                    if (targetTitle.ToString() == sourceTitle.ToString())
                    {
                        var sourceValue = sourceRangeValue[c, j];
                        if (sourceValue == null)
                        {
                            sourceValue = "";
                        }
                        var targetCell = targetSheet.Cells[targetCol, i + 1];
                        targetCell.Value = sourceValue;
                    }
                }
            }
            targetColList.Add(targetCol);
        }
        return targetColList;
    }

    public static List<(string, string, string)> SetExcelObjectEpPlus(dynamic excelPath,dynamic excelName,out ExcelWorksheet sheet  ,out ExcelPackage excel)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        sheet = null;
        excel = null;
        string errorExcelLog;
        var errorList = new List<(string, string, string)>();
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
        bool fileExists = File.Exists(path);
        if (fileExists == false)
        {
            errorExcelLog = excelName + "不存在表格文件";
            errorList.Add((excelName, errorExcelLog, excelName));
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
            errorExcelLog = excelName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelName, errorExcelLog, excelName));
            return errorList;
        }
        try
        {
            sheet = workBook.Worksheets["Sheet1"];
        }
        catch (Exception ex)
        {
            errorExcelLog = excelName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelName, errorExcelLog, excelName));
            return errorList;
        }
        sheet ??= workBook.Worksheets[0];
        return errorList;
    }

    //数字转Excel字母列
    public static string ConvertToExcelColumn(int columnNumber)
    {
        string columnName = "";

        while (columnNumber > 0)
        {
            int remainder = (columnNumber - 1) % 26;
            columnName = (char)('A' + remainder) + columnName;
            columnNumber = (columnNumber - 1) / 26;
        }
        return columnName;
    }
    //打开指定Excel文件，并定位到指定sheet的指定单元格
    public static void OpenExcelAndSelectCell(string filePath, string sheetName, string cellAddress)
    {
        try
        {
            // 打开指定路径的 Excel 文件
            var workbook = App.Workbooks.Open(filePath);
            // 获取指定名称的工作表
            var worksheet = workbook.Sheets[sheetName];
            // 选择指定的单元格
            var cellRange = worksheet.Range[cellAddress];
            cellRange.Select();
        }
        catch (Exception ex)
        {
            // 处理异常
            // ...
        }
        // 垃圾回收
        GC.Collect();
    }
    //EPPlus创建新Excel表格或获取已经存在的表格
    public static List<(string, string, string)> OpenOrCreatExcelByEpPlus(string excelFilePath,string excelName,out ExcelWorksheet sheet, out ExcelPackage excel)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        sheet = null;
        excel = null;
        var path = excelFilePath + @"\" + excelName+@".xlsx";
        if (!File.Exists(excelFilePath))
        {
            using (ExcelPackage packageCreat = new ExcelPackage())
            {
                var sheetCreat= packageCreat.Workbook.Worksheets.Add("Sheet1");
                FileInfo excelFile = new FileInfo(path);
                packageCreat.SaveAs(excelFile);
                sheetCreat.Dispose();
            }
        }
        var errorList = SetExcelObjectEpPlus(excelFilePath,excelName+@".xlsx",out  sheet, out  excel);
        return errorList;
    }


    ~PubMetToExcel()
    {
        App.Dispose();
    }

}