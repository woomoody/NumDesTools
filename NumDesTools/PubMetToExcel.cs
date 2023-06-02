using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using Color = System.Drawing.Color;

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

                            // 如果找到了匹配的值
                            if (cellValue != null && cellValue.ToString() == errorValue)
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
            currentCount++;
            //wk.Properties.Company = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
            App.StatusBar = "正在检查第" + currentCount + "/" + count + "个文件:" + file;
        }
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

}