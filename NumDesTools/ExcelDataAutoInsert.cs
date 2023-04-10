using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;



namespace NumDesTools;
public class ExcelDataAutoInsert
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    public static Dictionary<string, List<(string, string)>> ExcelModeIdDictionary;
    public static Dictionary<string, List<string>> ExcelFixKeyDictionary;
    public static Dictionary<string, List<string>> ExcelFixKeyMethodDictionary;
    public static List<string> ExcelFixGroup;

    public static int FindTitle(dynamic sheet,int rows,string findValue)
    {
        var maxColumn = sheet.UsedRange.Columns.Count;
        for (int column = 1; column <= maxColumn; column++)
        {
            var cell = sheet.Cells[rows, column] as Range;
            if(cell != null && cell.Value2?.ToString() == findValue)
            {
                return column;
            }
        }
        return -1;
    }
    public static int FindKeyColNpoi(string excelPath,string targetWorkbook,int rows,string findValue,string targetSheet="Sheet1")
    {

        var path = excelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet);
        if (sheet == null)
        {
            sheet = workbook.GetSheetAt(0);
        }
        var rowSource = sheet.GetRow(rows);
        for (int j = rowSource.FirstCellNum; j <= rowSource.LastCellNum; j++)
        {
            var cell = rowSource.GetCell(j);
            if (cell != null)
            {
                var cellValue = cell.ToString();
                if (cellValue == findValue)
                {
                    workbook.Close();
                    fs.Close();
                    return j;
                }
            }
        }
        workbook.Close();
        fs.Close();
        return 0;
    }
    public static int FindKeyRowNpoi(string excelPath, string targetWorkbook, int cols, string findValue, string targetSheet = "Sheet1")
    {
        var path = excelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet);
        if (sheet == null)
        {
            sheet = workbook.GetSheetAt(0);
        }
        for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
        {
            var rowSource = sheet.GetRow(i);
            if (rowSource != null)
            {
                var cell = rowSource.GetCell(cols);
                var cellValue = cell.ToString();
                if (cellValue == findValue)
                {
                    workbook.Close();
                    fs.Close();
                    return i;
                }
            }
        }
        workbook.Close();
        fs.Close();
        return -1;
    }
    public static int ExcelDic()
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        ExcelModeIdDictionary = new Dictionary<string, List<(string,string)>>();
        ExcelFixKeyDictionary = new Dictionary<string, List<string>>();
        ExcelFixKeyMethodDictionary = new Dictionary<string, List<string>>();
        ExcelFixGroup = new List<string>();
        var modeCol = FindTitle(sheet, 1, "模板");
        var excelCol = FindTitle(sheet, 1, "表名");
        var keyColFirst = FindTitle(sheet, 1, "修改字段");
        //读取模板表数据
        var rowsCount = sheet.UsedRange.Rows.Count;
        var colsCount = sheet.UsedRange.Columns.Count;
        for (var i = 2; i <= rowsCount; i++)
        {
            var cellExcel = sheet.Cells[i, excelCol].Value2;
            if (cellExcel == null) continue;
            var baseExcel = cellExcel.ToString();
            ExcelModeIdDictionary[baseExcel] = new List<(string, string)>();
            ExcelFixKeyDictionary[baseExcel] = new List<string>();
            ExcelFixKeyMethodDictionary[baseExcel] = new List<string>();
            ExcelFixGroup.Add(baseExcel);
            for (var j = keyColFirst; j <= colsCount; j++)
            {
                string baseExcelFixKey = sheet.Cells[i, j].Value2?.ToString();
                //var baseExcelFixKeyCol = FindKeyColNPOI(excelPath, baseExcel, 1, baseExcelFixKey);
                if (baseExcelFixKey == null)
                {
                    baseExcelFixKey = "";
                }
                ExcelFixKeyDictionary[baseExcel].Add(baseExcelFixKey);
                var baseExcelFixKeyMethod = sheet.Cells[i + 1, j].Value2;
                if (baseExcelFixKeyMethod == null)
                {
                    baseExcelFixKeyMethod = "";
                }
                ExcelFixKeyMethodDictionary[baseExcel].Add(baseExcelFixKeyMethod.ToString());
            }
            string baseExcelModeId1 = sheet.Cells[i, modeCol].Value2.ToString();
            string baseExcelModeId2 = sheet.Cells[i + 1, modeCol].Value2.ToString();
            var tuple = (baseExcelModeID1: baseExcelModeId1, baseExcelModeID2: baseExcelModeId2);
            if (string.IsNullOrEmpty(baseExcelModeId1) || string.IsNullOrEmpty(baseExcelModeId2))
            {
                MessageBox.Show(baseExcel+@":模板列第"+i+@"行有空值错误，不能导出");
                break;
            }
            ExcelModeIdDictionary[baseExcel].Add(tuple);
        }
        return excelCol;
    }
    public static void IgnoreExcel(string excelName)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var indexWk = App.ActiveWorkbook;
        var excelPath = indexWk.Path;
        var startValue = ExcelModeIdDictionary[excelName][0].Item1;
        var endValue = ExcelModeIdDictionary[excelName][0].Item2;
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
            case"UIItemConfigs.xlsx":
                path =newPath + @"\Excels\UIs\UIItemConfigs.xlsx";
                break;
            default: 
                path = excelPath + @"\" + excelName;
                break;
        }
        var excel = new ExcelPackage(new FileInfo(path));
        var workBook = excel.Workbook;
        var sheet = workBook.Worksheets[0];
        var startRowSource = ExcelRelationShipEpPlus. FindSourceRow(sheet, 2, startValue);
        var endRowSource = ExcelRelationShipEpPlus. FindSourceRow(sheet, 2, endValue);
        var colCount = sheet.Dimension.Columns;
        var count = endRowSource - startRowSource + 1;
        //数据复制
        sheet.InsertRow(endRowSource + 1, count);
        var cellSource = sheet.Cells[startRowSource, 2, endRowSource, colCount];
        var cellTarget = sheet.Cells[endRowSource + 1, 2, endRowSource + count, colCount];
        cellSource.Copy(cellTarget,ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting);
        //数据修改
        int countCc = 0;
        foreach (var keyIndex in ExcelFixKeyDictionary[excelName])
        {
            if(keyIndex == "") continue;
            //查找字段所在列
            string excelKey = ExcelFixKeyDictionary[excelName][countCc];
            int excelFileFixKey = FindKeyColNpoi(excelPath, excelName, 1, excelKey);
            //字典会把空值当0用
            if (excelFileFixKey == 0)
            {
                countCc++;
                continue;
            }
            //修改字段字典中的字段值，各自方法不一
            for (int i = 0; i < count; i++)
            {
                var cellFix = sheet.Cells[endRowSource + i + 1, excelFileFixKey + 1];
                if (cellFix.Value == null)
                {
                    continue;
                }
                //字段每个数字位数统计，原始modeID统计
                ExcelRelationShipEpPlus.KeyBitCount(cellFix.Value.ToString());
                //字段值改写方法
                var temp1 = ExcelRelationShipEpPlus.CellFixValueKeyList(ExcelFixKeyMethodDictionary[excelName][countCc]);
                //修改字符串
                var cellFixValue = StringRegPlace(cellFix.Value.ToString(), temp1, 1 + i);
                double number;
                if (double.TryParse(cellFixValue, out number))
                {
                    cellFix.Value = number;
                }
                else
                {
                    cellFix.Value = cellFixValue;
                }
            }
            countCc++;
            
        }
        excel.Save();
        excel.Dispose();
        App.StatusBar = "正在处理:" + excelName + "文件";

    }

    public static string AutoInsertDat()
    {
        var sw = new Stopwatch();
        sw.Start();

            var excelCol=ExcelDic();

        var ts1 = Math.Round(sw.Elapsed.TotalSeconds,2);
        var str1 = "字典用时:" + ts1;

            foreach (var excelName in ExcelFixGroup)
            {
                IgnoreExcel(excelName);
            }

        var ts2 = Math.Round(sw.Elapsed.TotalSeconds - ts1,2);
        var str2 = "写入数据用时:" + ts2;

        CellFormatAuto(excelCol);

        var ts3 = Math.Round(sw.Elapsed.TotalSeconds - ts2 - ts1,2);
        var str3 = "整理格式用时:" + ts3;

        ExcelHyperLinks();

        var ts4 = Math.Round(sw.Elapsed.TotalSeconds - ts2 - ts1 - ts3,2);
        var str4 = "构建超链接用时:" + ts4;

        var str = str1 +"#"+ str2 +"#"+ str3 + "#" + str4;
        return str;
    }
    public static string StringRegPlace(string str,List<(int,int)>digit,int addValue)
    {
        var reg = "\\d+";
        var matches = Regex.Matches(str, reg);
        var matchCount = 0;
        var digitCount = 0;
        foreach (Match unused in matches)
        {
            var matches2 = Regex.Matches(str, reg);
            var match2 = matches2[matchCount];
            var numStr = match2.Value;
            int index = match2.Index;
            var num = long.Parse(numStr);
            if (digit.Any(item => item.Item1 == matchCount+1))
            {
                //数字累加
                var newNum = num + (long)Math.Pow(10, digit[digitCount].Item2 - 1) * addValue;
                //字符替换
                var numCount = numStr.Length;
                str= str.Substring(0,index)+newNum+str.Substring(index+numCount);
                digitCount++;
            }
            else if (digit.Count == 1 && digit[0].Item1 == 0)
            {
                //数字累加
                var newNum = num + (long)Math.Pow(10, digit[0].Item2 - 1) * addValue;
                //字符替换
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
            }
            matchCount++;
        }
        return str;
    }
    public static void ExcelHyperLinks()
    {
        var indexWk = App.ActiveWorkbook;
        var excelPath = indexWk.Path;
        var sheet = indexWk.ActiveSheet;
        for (var i = 2; i <= 500; i++)
        {
            //找到模板表所在行
            var modeCol = FindTitle(sheet, 1, "模板");
            var excelName = FindTitle(sheet, 1, "表名");
            string findValue = sheet.Cells[i + 1, modeCol].Value?.ToString();
            var cell = sheet.Cells[i, excelName];
            if (cell.value != null && cell.value.ToString().Contains(".xlsx"))
            {
                int row = FindKeyRowNpoi(excelPath, cell.value.ToString(), 1, findValue)+1;
                if (row != 0)
                {
                    var newRow = "A" + row;
                    var excel = new FileStream(excelPath + @"\" + cell.value.ToString(), FileMode.Open, FileAccess.Read);
                    var workbook = new XSSFWorkbook(excel);
                    var sheetTemp = workbook.GetSheetAt(0);
                    var sheetName = sheetTemp.SheetName;
                    var links = excelPath + @"\" + cell.value.ToString() + "#" + sheetName + "!" + newRow;
                    workbook.Close();
                    excel.Close();
                    cell.Hyperlinks.Add(cell, links);
                    cell.Font.Size = 9;
                    cell.Font.Name = "微软雅黑";
                }
            }
        }
    }
    public static void CellFormatAuto(int excelCol)
    {
        var indexWk = App.ActiveWorkbook;
        var sheet = indexWk.ActiveSheet;
        var rowsCount = (sheet.Cells[sheet.Rows.Count, "B"].End[XlDirection.xlUp].Row - 1) / 2 + 1;
        for (var i = 1; i <= rowsCount; i++)
        {
            for (var j = 0; j <= 20; j++)
            {
                var cell = sheet.Cells[1, 1].Offset [(i - 1) * 2+1, j];
                var cell2 = sheet.Cells[1, 1].Offset[(i - 1) * 2+2, j];
                if (cell.Value != null)
                {
                    cell2.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDashDotDot;
                    cell2.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                }
                else
                {
                    cell2.Borders.LineStyle = XlLineStyle.xlDash;
                    cell2.Borders.Weight = XlBorderWeight.xlHairline;
                }
            }
            var c1 = sheet.Cells[1, 1].Offset[(i - 1) * 2 + 1, excelCol-1];
            var c2 = sheet.Cells[1, 1].Offset[(i - 1) * 2 + 2, excelCol-1];
            var rng = sheet.Range[c1, c2];
            rng.Merge();
        }
    }
}