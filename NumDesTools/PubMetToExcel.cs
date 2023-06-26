using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;
using Point = Microsoft.Office.Interop.Excel.Point;
using System.Threading.Tasks;

namespace NumDesTools;

public class PubMetToExcel
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToList(dynamic workSheet)
    {
        Range dataRange = workSheet.UsedRange;
        //// 获取第一列有数据的最大行
        //int maxRow= 0;
        //for (int row = 1; row <= dataRange.Rows.Count; row++)
        //{
        //    if (workSheet.Cells[row, 1].Value != null && workSheet.Cells[row, 1].Value.ToString().Trim() != "")
        //    {
        //        maxRow = row;
        //    }
        //}
        //Range realDataRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[maxRow, dataRange.Columns.Count]];
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
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToListFirstRow(dynamic workSheet)
    {
        Range dataRange = workSheet.UsedRange;
        // 获取第一列有数据的最大行
        int maxRow = 0;
        for (int row = 1; row <= dataRange.Rows.Count; row++)
        {
            if (workSheet.Cells[row, 1].Value != null && workSheet.Cells[row, 1].Value.ToString().Trim() != "")
            {
                maxRow = row;
            }
        }
        Range realDataRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[maxRow, dataRange.Columns.Count]];
        // 读取数据到一个二维数组中
        object[,] rangeValue = realDataRange.Value;
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
            if (dic.ContainsKey(value))
            {
                dic[value].Add(tuple);
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
    /*
        public static (string file, string Name, int cellRow, int cellCol) ErrorKeyFromExcelMulti(string rootPath, string errorValue)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(rootPath));
            var mainPath = newPath + @"\Excels\Tables\";
            string[] files1 = Directory.EnumerateFiles(mainPath, "*.xlsx")
                .Where(file => !Path.GetFileName(file).Contains("#"))
                .ToArray();
            var langPath = newPath + @"\Excels\Localizations\";
            string[] files2 = Directory.EnumerateFiles(mainPath, "*.xlsx")
                .Where(file => !Path.GetFileName(file).Contains("#"))
                .ToArray();
            var uiPath = newPath + @"\Excels\UIs\";
            string[] files3 = Directory.EnumerateFiles(mainPath, "*.xlsx")
                .Where(file => !Path.GetFileName(file).Contains("#"))
                .ToArray();
            var files = files1.Concat(files2).Concat(files3).ToArray();

            int currentCount = 0;
            var count = files.Length;

            // 定义一个共享变量用于存储匹配结果
            (string, string, int, int) result = ("", "", 0, 0);

            // 创建一个 object 类型的对象用于锁定
            object locker = new object();

            Parallel.For(0, files.Length, (i) =>
            {
                //过滤非配置表
                var fileName = Path.GetFileName(files[i]);
                if (fileName.Contains("#")) return;
                // 使用 EPPlus 打开 Excel 文件进行操作
                using (ExcelPackage package = new ExcelPackage(new FileInfo(files[i])))
                {
                    var wk = package.Workbook;
                    var sheet = wk.Worksheets["Sheet1"] ?? wk.Worksheets[0];
                    for (var col = 2; col <= sheet.Dimension.End.Column; col++)
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
                                var tuple = (files[i], sheet.Name, cellRow, cellCol);

                                // 在锁定的情况下更新共享变量
                                lock (locker)
                                {
                                    result = tuple;
                                    // 更新计数器
                                    currentCount++;
                                    // 更新状态栏
                                    App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + files[i];
                                }

                                // 跳出循环并返回结果
                                return;
                            }
                        }
                    }
                }

                // 更新计数器
                lock (locker)
                {
                    currentCount++;
                    // 更新状态栏
                    App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + files[i];
                }
            });

            return result;
        }
    */
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
                                if (cellValue != null && cellValue.ToString()==errorValue)
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
    public static List<int> MergeExcel(object[,] sourceRangeValue, ExcelWorksheet targetSheet, object[,] targetRangeTitle,
        object[,] sourceRangeTitle)
    {
        var targetRowList =new List<int>();
        int defaultRow = targetSheet.Dimension.End.Row; ;
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
            var taregtRow = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, sourceRow.ToString());
            if (taregtRow == -1)
            {
                targetSheet.InsertRow(beforTargetRow + 1,1);
                taregtRow = beforTargetRow + 1;
            }
            else
            {
                beforTargetRow = taregtRow;
            }
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
                        var targetCell = targetSheet.Cells[taregtRow, i + 1];
                        targetCell.Value = sourceValue;
                    }
                }
            }
            targetRowList.Add(taregtRow);
        }
        return targetRowList;
    }

    public static List<(string, string, string)> EpplusCreatExcelObj(dynamic excelPath,dynamic excelName,out ExcelWorksheet sheet  ,out ExcelPackage excel)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        sheet = null;
        excel = null;
        var errorExcelLog = "";
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

    public static DialogResult ExMessageBox(string message,string filePath)
    {
        var f = new DataExportForm
        {
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(400, 200),
            MaximizeBox = false,
            MinimizeBox = false,
            Text = @"表格汇总"
        };
        var gb = new Panel
        {
            BackColor = Color.FromArgb(255, 225, 225, 225),
            AutoScroll = true,
            Location = new System.Drawing.Point(f.Left + 20, f.Top + 20),
            Size = new Size(f.Width - 55, f.Height - 200),
            Text = message
        };
        //gb.Dock = DockStyle.Fill;
        f.Controls.Add(gb);
        var bt3 = new Button
        {
            Name = "button3",
            Text = @"导出",
            Location = new System.Drawing.Point(f.Left + 360, f.Top + 680)
        };
        f.Controls.Add(bt3);
        return f.ShowDialog();
        bt3.Click += Btn3Click;
        void Btn3Click(object sender, EventArgs e)
        {
            Process.Start(filePath);
        }
    }
}