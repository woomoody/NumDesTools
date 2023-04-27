using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDna.Integration;

namespace NumDesTools;

public class PubMetToExcel
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
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

    public static void RepeatValue(ExcelWorksheet sheet,int row, int col,string repeatValue)
    {
        // 获取指定列的单元格数据
        var colA = sheet.Cells[row, col, sheet.Dimension.End.Row, col];

        // 通过LINQ查询获取重复值
        var duplicates = colA
            .Where(cell => cell.Value.ToString().Equals(repeatValue))
            .GroupBy(cell => cell.Value.ToString())
            .Where(group => group.Count() > 1)
            .SelectMany(group => group.Skip(1));
        // 判断是否存在指定的重复值
        var duplicateCells = duplicates.ToList();
        if (duplicateCells.Any())
        {
            // 删除重复值所在的行
            foreach (var duplicateCell in duplicateCells)
            {
                sheet.DeleteRow(duplicateCell.Start.Row);
            }
        }
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
    public static (string file, string Name, int cellRow, int cellCol) ErrorKeyFromExcelAll(string rootPath, string errorValue)
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
            // 使用 EPPlus 打开 Excel 文件进行操作
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
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
}