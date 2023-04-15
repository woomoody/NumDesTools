 using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;
 using System.ComponentModel;
 using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
using Color = System.Drawing.Color;
using OfficeOpenXml.DataValidation;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows;
 using MessageBox = System.Windows.Forms.MessageBox;
using OfficeOpenXml.Style;
using NPOI.SS.Formula.Functions;
 using Org.BouncyCastle.Asn1.Nist;
 using LicenseContext = OfficeOpenXml.LicenseContext;
 using Match = System.Text.RegularExpressions.Match;

 namespace NumDesTools;
public class ExcelDataAutoInsert
{
    public static int FindTitle(dynamic sheet,int rows,string findValue)
    {
        var maxColumn = sheet.UsedRange.Columns.Count;
        for (int column = 1; column <= maxColumn; column++)
        {
            if(sheet.Cells[rows, column] is Range cell && cell.Value2?.ToString() == findValue)
            {
                return column;
            }
        }
        return -1;
    }
    public static int FindSourceCol(ExcelWorksheet sheet, int row, string searchValue)
    {
        for (int col = 2; col <= sheet.Dimension.End.Column; col++)
        {
            // 获取当前行的单元格数据
            var cellValue = sheet.Cells[row, col].Value;

            // 如果找到了匹配的值
            if (cellValue != null && cellValue.ToString() == searchValue)
            {
                // 返回该单元格的行地址
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Column;
                return rowAddress;
            }
        }
        return -1;
    }
    public static int FindSourceRow(ExcelWorksheet sheet, int col, string searchValue)
    {
        for (int row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            // 获取当前行的单元格数据
            var cellValue = sheet.Cells[row, col].Value;

            // 如果找到了匹配的值
            if (cellValue != null && cellValue.ToString() == searchValue)
            {
                // 返回该单元格的行地址
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Row;
                return rowAddress;
            }
        }
        return -1;
    }
    //public static int FindKeyColNpoi(string excelPath,string targetWorkbook,int rows,string findValue,string targetSheet="Sheet1")
    //{
    //    string path;
    //    var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
    //    switch (targetWorkbook)
    //    {
    //        case "Localizations.xlsx":
    //            path = newPath + @"\Excels\Localizations\Localizations.xlsx";
    //            break;
    //        case "UIConfigs.xlsx":
    //            path = newPath + @"\Excels\UIs\UIConfigs.xlsx";
    //            break;
    //        case "UIItemConfigs.xlsx":
    //            path = newPath + @"\Excels\UIs\UIItemConfigs.xlsx";
    //            break;
    //        default:
    //            path = excelPath + @"\" + targetWorkbook;
    //            break;
    //    }
    //    var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    //    var workbook = new XSSFWorkbook(fs);
    //    var sheet = workbook.GetSheet(targetSheet);
    //    if (sheet == null)
    //    {
    //        sheet = workbook.GetSheetAt(0);
    //    }
    //    var rowSource = sheet.GetRow(rows);
    //    for (int j = rowSource.FirstCellNum; j <= rowSource.LastCellNum; j++)
    //    {
    //        var cell = rowSource.GetCell(j);
    //        if (cell != null)
    //        {
    //            var cellValue = cell.ToString();
    //            if (cellValue == findValue)
    //            {
    //                workbook.Close();
    //                fs.Close();
    //                return j;
    //            }
    //        }
    //    }
    //    workbook.Close();
    //    fs.Close();
    //    return 0;
    //}
    //public static int FindKeyRowNpoi(string excelPath, string targetWorkbook, int cols, string findValue, string targetSheet = "Sheet1")
    //{
    //    string path;
    //    var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
    //    switch (targetWorkbook)
    //    {
    //        case "Localizations.xlsx":
    //            path = newPath + @"\Excels\Localizations\Localizations.xlsx";
    //            break;
    //        case "UIConfigs.xlsx":
    //            path = newPath + @"\Excels\UIs\UIConfigs.xlsx";
    //            break;
    //        case "UIItemConfigs.xlsx":
    //            path = newPath + @"\Excels\UIs\UIItemConfigs.xlsx";
    //            break;
    //        default:
    //            path = excelPath + @"\" + targetWorkbook;
    //            break;
    //    }
    //    var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    //    var workbook = new XSSFWorkbook(fs);
    //    var sheet = workbook.GetSheet(targetSheet);
    //    if (sheet == null)
    //    {
    //        sheet = workbook.GetSheetAt(0);
    //    }
    //    for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
    //    {
    //        var rowSource = sheet.GetRow(i);
    //        if (rowSource != null)
    //        {
    //            var cell = rowSource.GetCell(cols);
    //            var cellValue = cell.ToString();
    //            if (cellValue == findValue)
    //            {
    //                workbook.Close();
    //                fs.Close();
    //                return i;
    //            }
    //        }
    //    }
    //    workbook.Close();
    //    fs.Close();
    //    return -1;
    //}
    private static List<(int, int)> ExcelDic(dynamic ExcelModeIdDictionary, dynamic ExcelModeIdNewDictionary, dynamic ExcelFixKeyDictionary, dynamic ExcelFixKeyMethodDictionary, dynamic ExcelFixGroup,dynamic sheet)
    {
        var modeCol = FindTitle(sheet, 1, "初始模板");
        var modeColNew = FindTitle(sheet, 1, "实际模板(上一期)");
        var excelCol = FindTitle(sheet, 1, "表名");
        var keyColFirst = FindTitle(sheet, 1, "修改字段");
        var addValueIndexMax = FindTitle(sheet, 1, "创建期号");
        var addValueIndexMin = FindTitle(sheet, 1, "模板期号");
        var addValue = Convert.ToInt32(sheet.Cells[2, addValueIndexMax].Value- sheet.Cells[2, addValueIndexMin].Value);
        var defaultData = new List<(int,int)>();
        defaultData.Add((excelCol,addValue));
        //读取模板表数据
        var rowsCount = sheet.UsedRange.Rows.Count;
        var colsCount = sheet.UsedRange.Columns.Count;
        int excelCount=0;
        for (var i = 2; i <= rowsCount; i++)
        {
            var cellExcel = sheet.Cells[i, excelCol].Value2;
            if (cellExcel == null) continue;
            var baseExcel = cellExcel.ToString();
            ExcelModeIdDictionary[excelCount] = new List<(string, string)>();
            ExcelModeIdNewDictionary[excelCount] =new List<(string, string)>();
            ExcelFixKeyDictionary[excelCount] = new List<string>();
            ExcelFixKeyMethodDictionary[excelCount] = new List<string>();
            ExcelFixGroup.Add(baseExcel);
            for (var j = keyColFirst; j <= colsCount; j++)
            {
                string baseExcelFixKey = sheet.Cells[i, j].Value2?.ToString();
                //var baseExcelFixKeyCol = FindKeyColNPOI(excelPath, baseExcel, 1, baseExcelFixKey);
                if (baseExcelFixKey == null)
                {
                    baseExcelFixKey = "";
                }
                ExcelFixKeyDictionary[excelCount].Add(baseExcelFixKey);
                var baseExcelFixKeyMethod = sheet.Cells[i + 1, j].Value2;
                if (baseExcelFixKeyMethod == null)
                {
                    baseExcelFixKeyMethod = "";
                }
                ExcelFixKeyMethodDictionary[excelCount].Add(baseExcelFixKeyMethod.ToString());
            }
            string baseExcelModeId1 = sheet.Cells[i, modeCol].Value2.ToString();
            string baseExcelModeId2 = sheet.Cells[i + 1, modeCol].Value2.ToString();
            string baseExcelModeId3 = sheet.Cells[i, modeColNew].Value2.ToString();
            string baseExcelModeId4 = sheet.Cells[i + 1, modeColNew].Value2.ToString();
            var tuple = (baseExcelModeID1: baseExcelModeId1, baseExcelModeID2: baseExcelModeId2);
            var tuple2 = (baseExcelModeID1: baseExcelModeId3, baseExcelModeID2: baseExcelModeId4);
            if (string.IsNullOrEmpty(baseExcelModeId1) || string.IsNullOrEmpty(baseExcelModeId2))
            {
                MessageBox.Show(baseExcel+@":模板列第"+i+@"行有空值错误，不能导出");
                break;
            }
            ExcelModeIdDictionary[excelCount].Add(tuple);
            ExcelModeIdNewDictionary[excelCount].Add(tuple2);
            excelCount++;
        }
        return defaultData;
    }
    private static List<(int, string, string)> SingleExcelDataWrite(int excelCount,int addValue,dynamic ExcelFixGroup, dynamic ExcelModeIdDictionary, dynamic ExcelModeIdDNewictionary, dynamic ExcelFixKeyDictionary, dynamic ExcelFixKeyMethodDictionary, dynamic excelPath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var excelName = ExcelFixGroup[excelCount];
        var startValue = ExcelModeIdDictionary[excelCount][0].Item1;
        var endValue = ExcelModeIdDictionary[excelCount][0].Item2;
        var isInertRowValue = ExcelModeIdDNewictionary[excelCount][0].Item1;
        int errorExcel=0;
        string errorExcelLog="";
        var errorList = new List<(int,string,string)>();
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
        ExcelWorkbook workBook = null;
        try
        {
            workBook = excel.Workbook;
        }
        catch(Exception ex)
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((errorExcel, errorExcelLog,excelName));
            return errorList;
        }
        ExcelWorksheet sheet = null;
        try
        {
            sheet = workBook.Worksheets["Sheet1"];
        }
        catch (Exception ex)
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((errorExcel, errorExcelLog,excelName));
            return errorList;
        }
        if (sheet == null)
        {
            sheet = workBook.Worksheets[0];
        }
        var startRowSource =  FindSourceRow(sheet, 2, startValue);
        if (startRowSource == -1)
        {
            errorExcel=excelCount * 2 + 2;
            errorExcelLog=excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
            errorList.Add((errorExcel,errorExcelLog, excelName));
            return errorList;
        }
        var endRowSource =  FindSourceRow(sheet, 2, endValue);
        if (endRowSource == -1)
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
            errorList.Add((errorExcel, errorExcelLog, excelName));
            return errorList;
        }
        var colCount = sheet.Dimension.Columns;
        var count = endRowSource - startRowSource + 1;
        //数据复制
        var isInertRowTarget = FindSourceRow(sheet, 2, isInertRowValue);
        if (isInertRowValue != "")
        {

            if (isInertRowTarget == -1)
            {
                sheet.InsertRow(endRowSource + 1, count);
                var cellSource = sheet.Cells[startRowSource, 1, endRowSource, colCount];
                var cellTarget = sheet.Cells[endRowSource + 1, 1, endRowSource + count, colCount];
                cellSource.Copy(cellTarget, ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting | ExcelRangeCopyOptionFlags.ExcludeMergedCells);
                cellSource.CopyStyles(cellTarget);
            }
        }
        else
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#【实际模板（上一期）】#[" + isInertRowValue + "]未找到(序号出错)";
            errorList.Add((errorExcel, errorExcelLog,excelName));
            return errorList;
        }
        //数据修改
        int countCc = 0;
        foreach (var keyIndex in ExcelFixKeyDictionary[excelCount])
        {
            if (keyIndex == "") continue;
            //查找字段所在列
            string excelKey = ExcelFixKeyDictionary[excelCount][countCc];
            int excelFileFixKey = FindSourceCol(sheet, 2, excelKey);
            //字典会把空值当0用
            if (excelFileFixKey == -1)
            {
                countCc++;
                continue;
            }
            //修改字段字典中的字段值，各自方法不一
            for (int i = 0; i < count; i++)
            {
                var cellSource = sheet.Cells[startRowSource + i, excelFileFixKey];
                var cellFix = sheet.Cells[endRowSource + i + 1, excelFileFixKey];
                if (cellSource.Value == null)
                {
                    continue;
                }
                if (cellSource.Value.ToString() == "")
                {
                    continue;
                }
                //字段每个数字位数统计，原始modeID统计
                //KeyBitCount(cellFix.Value.ToString());
                //字段值改写方法
                var temp1 = CellFixValueKeyList(ExcelFixKeyMethodDictionary[excelCount][countCc]);
                //修改字符串
                var cellFixValue = StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                if (cellFixValue == "^error^")
                {
                    errorExcel = excelCount * 2 + 2;
                    errorExcelLog = excelName + "#【修改模式】#[" + excelKey + "]字段方法写错";
                    errorList.Add((errorExcel, errorExcelLog, excelName));
                    return errorList;
                }
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
        errorList.Add((errorExcel, errorExcelLog, excelName));
        return errorList;
    }

    public static string AutoInsertDat(bool threadMode)
    {
        dynamic App = ExcelDnaUtil.Application;
        Dictionary<int, List<(string, string)>> ExcelModeIdDictionary =new Dictionary<int, List<(string, string)>>();
        Dictionary<int, List<(string, string)>> ExcelModeIdNewDictionary = new Dictionary<int, List<(string, string)>>();
        Dictionary<int, List<string>> ExcelFixKeyDictionary =new Dictionary<int, List<string>>();
        Dictionary<int, List<string>> ExcelFixKeyMethodDictionary =new Dictionary<int, List<string>>();
        List<string> ExcelFixGroup = new List<string>();
        dynamic indexWk = App.ActiveWorkbook;
        dynamic sheet = indexWk.ActiveSheet;
        var excelPath = indexWk.Path;

        ErrorLogCtp.DisposeCtp();

        var sw = new Stopwatch();
        sw.Start();
        //获取字典
        var defaultData=ExcelDic(ExcelModeIdDictionary,ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, ExcelFixGroup,sheet);
        var ts1 = Math.Round(sw.Elapsed.TotalSeconds,2);
        var str1 = "字典用时:" + ts1;
        //遍历文件
        var excelCount = 0;
         var errorExcelList = new List<List<(int,string,string)>>();
        foreach (var excelName in ExcelFixGroup)
        {
            List<(int, string,string)> error = null;
            if (threadMode)
            {
                error = MultiExcelDataWrite(excelCount, defaultData[0].Item2, ExcelFixGroup, ExcelModeIdDictionary, ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, excelPath);
            }
            else
            {
                error = SingleExcelDataWrite(excelCount, defaultData[0].Item2, ExcelFixGroup, ExcelModeIdDictionary, ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, excelPath);
            }
            if (error.Count != 0)
            {
                errorExcelList.Add(error);
            }
            excelCount++;
            App.StatusBar = "写入数据" + "<" + excelCount + "/" + ExcelFixGroup.Count + ">" + excelName;
        }
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds - ts1,2);
        var str2 = "写入数据用时:" + ts2;
        //出错表格处理
        string errorLog = ErrorExcelMark(errorExcelList,sheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }
        else
        {
            sheet.Range["A3:A1000"].Value = "";
        }
        //CellFormatAuto(defaultData[0].Item1);
        //var ts3 = Math.Round(sw.Elapsed.TotalSeconds - ts2 - ts1,2);
        //var str3 = "整理格式用时:" + ts3;
        //ExcelHyperLinks();
        //var ts4 = Math.Round(sw.Elapsed.TotalSeconds - ts2 - ts1 - ts3,2);
        //var str4 = "构建超链接用时:" + ts4;
        var str = str1 + "#" + str2;//+ "#" + str3;//+ "#"+ str4;
        return str;
    }

    public static string ErrorExcelMark(dynamic errorExcelList,dynamic sheet )
    {
        var strBuild = new StringBuilder();
        for (int i = 0;i< errorExcelList.Count;i++)
        {
            if (errorExcelList[i][0].Item1 == 0) continue;
            strBuild.Append(errorExcelList[i][0].Item2);
            var cell = sheet.Cells[errorExcelList[i][0].Item1, 1];
            cell.Value = "git checkout -- Excels/Tables/"+ errorExcelList[i][0].Item3;
            cell.Font.Color = Color.Red;
            //cell.Font.Size = 9;
            //cell.Font.Name = "微软雅黑";
            //cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //cell.VerticalAlignment = XlHAlign.xlHAlignCenter;
            //cell.ColumnWidth = 8.38;
            //cell.RowHeight = 14.25;
            //cell.ShrinkToFit = true;
            //cell.Borders.LineStyle = XlLineStyle.xlDash;
            //cell.Borders.Weight = XlBorderWeight.xlHairline;
        }
        var errorLog = strBuild.ToString();
        return errorLog;
    }
    private static string StringRegPlace(string str,List<(int,int)>digit,int addValue)
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
                var addDigit = (long)Math.Pow(10, digit[digitCount].Item2 - 1) * addValue;
                if (addDigit >= num)
                {
                    str = "^error^";
                    return str;
                }
                var newNum = num + addDigit;
                //字符替换
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
                digitCount++;
            }
            else if (digit.Count == 1 && digit[0].Item1 == 0)
            {
                //数字累加
                var addDigit =Math.Abs( (long)Math.Pow(10, digit[0].Item2 - 1) * addValue);
                if (addDigit >= num)
                {
                    str = "^error^";
                    return str;
                }
                var newNum = num + addDigit;
                //字符替换
                var numCount = numStr.Length;
                str = str.Substring(0, index) + newNum + str.Substring(index + numCount);
            }
            matchCount++;
        }
        return str;
    }
    public static void ExcelHyperLinks(dynamic excelPath, dynamic sheet)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        for (var i = 2; i <= 500; i++)
        {
            //找到模板表所在行
            var modeCol = FindTitle(sheet, 1, "初始模板");
            var excelName = FindTitle(sheet, 1, "表名");
            string findValue = sheet.Cells[i , modeCol].Value?.ToString();
            var cell = sheet.Cells[i, excelName];
            string path;
            if (cell.value != null && cell.value.ToString().Contains(".xlsx"))
            {
                var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
                switch (cell.value)
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
                        path = excelPath + @"\" + cell.value;
                        break;
                }
                var excel = new ExcelPackage(new FileInfo(path));
                var workbook = excel.Workbook;
                var sheetTemp = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];
                int row = FindSourceRow(sheetTemp, 2, findValue);
                if (row != 0)
                {
                    var newRow = "A" + row;
    
                    var sheetName = sheetTemp.Name;
                    var links = path + "#" + sheetName + "!" + newRow;
                    excel.Dispose();
                    cell.Hyperlinks.Add(cell, links);
                    cell.Font.Size = 9;
                    cell.Font.Name = "微软雅黑";
                }
            }
        }
    }
    public static void CellFormatAuto(dynamic ExcelModeIdDictionary, dynamic ExcelModeIdNewDictionary, dynamic ExcelFixKeyDictionary, dynamic ExcelFixKeyMethodDictionary, dynamic ExcelFixGroup, dynamic sheet)
    {
        var defaultData = ExcelDic(ExcelModeIdDictionary, ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, ExcelFixGroup, sheet);
        var excelCol = defaultData[0].Item1;
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
    public static void RightClickWriteExcel(CommandBarButton ctrl, ref bool cancelDefault)
    {
        dynamic App = ExcelDnaUtil.Application;
        Dictionary<int, List<(string, string)>> ExcelModeIdDictionary = new Dictionary<int, List<(string, string)>>();
        Dictionary<int, List<(string, string)>> ExcelModeIdNewDictionary = new Dictionary<int, List<(string, string)>>();
        Dictionary<int, List<string>> ExcelFixKeyDictionary = new Dictionary<int, List<string>>();
        Dictionary<int, List<string>> ExcelFixKeyMethodDictionary = new Dictionary<int, List<string>>();
        List<string> ExcelFixGroup = new List<string>();
        dynamic indexWk = App.ActiveWorkbook;
        dynamic sheet = indexWk.ActiveSheet;
        var excelPath = indexWk.Path;

        ErrorLogCtp.DisposeCtp();

        var sw = new Stopwatch();
        sw.Start();
        var defaultData = ExcelDic(ExcelModeIdDictionary, ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, ExcelFixGroup, sheet);
        var errorExcelList = new List<List<(int, string,string)>>();

        var cell = App.Selection;
        var rowStart = cell.Row;
        var rowCount =cell.Rows.Count;
        var rowEnd = rowStart + rowCount - 1;
        for (int i = rowStart; i <= rowEnd; i++)
        {
            var realExcel =(i-2) % 2;
            if (realExcel == 0)
            {
                int excelCount = (i - 2) / 2;
                var error = SingleExcelDataWrite(excelCount, defaultData[0].Item2, ExcelFixGroup, ExcelModeIdDictionary,
                    ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, excelPath);
                errorExcelList.Add(error);
                App.StatusBar = "写入数据" + "<" + excelCount + "/" + ExcelFixGroup.Count + ">" + ExcelFixGroup[excelCount];
            }
        }
        //出错表格处理
        string errorLog = ErrorExcelMark(errorExcelList, sheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }
        else
        {
            sheet.Range["A3: A1000"].Value = "";
        }
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        var str2 = "写入数据用时:" + ts2;
        App.StatusBar = str2;
    }
    public static void RightClickWriteExcelThread(CommandBarButton ctrl, ref bool cancelDefault)
    {
        dynamic App = ExcelDnaUtil.Application;
        Dictionary<int, List<(string, string)>> ExcelModeIdDictionary = new Dictionary<int, List<(string, string)>>();
        Dictionary<int, List<(string, string)>> ExcelModeIdNewDictionary = new Dictionary<int, List<(string, string)>>();
        Dictionary<int, List<string>> ExcelFixKeyDictionary = new Dictionary<int, List<string>>();
        Dictionary<int, List<string>> ExcelFixKeyMethodDictionary = new Dictionary<int, List<string>>();
        List<string> ExcelFixGroup = new List<string>();
        dynamic indexWk = App.ActiveWorkbook;
        dynamic sheet = indexWk.ActiveSheet;
        var excelPath = indexWk.Path;

        ErrorLogCtp.DisposeCtp();

        var sw = new Stopwatch();
        sw.Start();
        var defaultData = ExcelDic(ExcelModeIdDictionary, ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, ExcelFixGroup, sheet);
        var errorExcelList = new List<List<(int, string,string)>>();

        var cell = App.Selection;
        var rowStart = cell.Row;
        var rowCount = cell.Rows.Count;
        var rowEnd = rowStart + rowCount - 1;
        for (int i = rowStart; i <= rowEnd; i++)
        {
            var realExcel = (i - 2) % 2;
            if (realExcel == 0)
            {
                int excelCount = (i - 2) / 2;
                var error = MultiExcelDataWrite(excelCount, defaultData[0].Item2, ExcelFixGroup, ExcelModeIdDictionary,
                    ExcelModeIdNewDictionary, ExcelFixKeyDictionary, ExcelFixKeyMethodDictionary, excelPath);

                errorExcelList.Add(error);
                App.StatusBar = "写入数据" + "<" + excelCount + "/" + ExcelFixGroup.Count + ">" + ExcelFixGroup[excelCount];
            }
        }
        //出错表格处理
        string errorLog = ErrorExcelMark(errorExcelList, sheet);
        if (errorLog != "")
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(errorLog);
        }
        else
        {
            sheet.Range["A3: A1000"].Value = "";
        }
        var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        var str2 = "写入数据用时:" + ts2;
        App.StatusBar = str2;
    }
    private static List<(int digitCount, string temp)> KeyBitCount(string str)
    {
        var regex = new Regex(@"\d+");
        var matches = regex.Matches(str);
        var keyBitCount = new List<(int digitCount, string temp)>();
        foreach (var match in matches)
        {
            var temp = match.ToString();
            var digitCount = temp.Length;
            keyBitCount.Add((digitCount, temp));
        }

        return keyBitCount;
    }
    private static List<(int, int)> CellFixValueKeyList(string str)
    {
        var monkeyList = new List<(int, int)>();

        str ??= "";

        if (str.Contains(','))
        {
            var pairs = str.Split(',');
            foreach (var pair in pairs)
            {
                if (pair.Contains('#'))
                {
                    var parts = pair.Split('#');
                    if (!int.TryParse(parts[0], out var key)) // 尝试将值解析为整数，如果解析失败就将值设为 0
                    {
                        MessageBox.Show($@"{str}#前必须有数值");
                        Environment.Exit(0);
                    }

                    if (!int.TryParse(parts[1], out var value)) // 尝试将值解析为整数，如果解析失败就将值设为 0
                        value = 1;

                    monkeyList.Add((key, value));
                }
                else
                {
                    monkeyList.Add((int.Parse(pair), 1));
                }
            }
        }
        else
        {
            if (str.Contains('#'))
            {
                var parts = str.Split('#');
                var key = Convert.ToInt32(parts[0]);
                var value = Convert.ToInt32(parts[1]);
                monkeyList.Add((key, value));
            }
            else
            {
                int strTemp;
                if (str == "")
                {
                    strTemp = 0;
                    monkeyList.Add((strTemp, 1));
                }
                else
                {
                    strTemp = int.Parse(str);
                    monkeyList.Add((0, strTemp));
                }
            }
        }

        return monkeyList;
    }
    private static List<(int, string, string)> MultiExcelDataWrite(int excelCount, int addValue, dynamic ExcelFixGroup, dynamic ExcelModeIdDictionary, dynamic ExcelModeIdDNewictionary, dynamic ExcelFixKeyDictionary, dynamic ExcelFixKeyMethodDictionary, dynamic excelPath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var excelName = ExcelFixGroup[excelCount];
        var startValue = ExcelModeIdDictionary[excelCount][0].Item1;
        var endValue = ExcelModeIdDictionary[excelCount][0].Item2;
        var isInertRowValue = ExcelModeIdDNewictionary[excelCount][0].Item1;
        int errorExcel =0;
        string errorExcelLog="";
        var errorList = new List<(int, string,string)>();
        string path = ExcelPathIgnore(excelPath, excelName);
        var excel = new ExcelPackage(new FileInfo(path));
        ExcelWorkbook workBook = null;
        try
        {
            workBook = excel.Workbook;
        }
        catch (Exception ex)
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((errorExcel, errorExcelLog, excelName));
            return errorList;
        }

        ExcelWorksheet sheet = null;
        try
        {
            sheet = workBook.Worksheets["Sheet1"];
        }
        catch (Exception ex)
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((errorExcel, errorExcelLog, excelName));
            return errorList;
        }
        if (sheet == null)
        {
            sheet = workBook.Worksheets[0];
        }
        var startRowSource = FindSourceRow(sheet, 2, startValue);
        if (startRowSource == -1)
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
            errorList.Add((errorExcel, errorExcelLog, excelName));
            return errorList;
        }
        var endRowSource = FindSourceRow(sheet, 2, endValue);
        if (endRowSource == -1)
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]未找到(序号出错)";
            errorList.Add((errorExcel, errorExcelLog, excelName));
            return errorList;
        }
        var colCount = sheet.Dimension.Columns;
        var count = endRowSource - startRowSource + 1;
        //数据复制
        var isInertRowTarget = FindSourceRow(sheet, 2, isInertRowValue);
        if (isInertRowValue != "")
        {
            if (isInertRowTarget == -1)
            {
                sheet.InsertRow(endRowSource + 1, count);
                var cellSource = sheet.Cells[startRowSource, 1, endRowSource, colCount];
                var cellTarget = sheet.Cells[endRowSource + 1, 1, endRowSource + count, colCount];
                cellSource.Copy(cellTarget, ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting | ExcelRangeCopyOptionFlags.ExcludeMergedCells);
                cellSource.CopyStyles(cellTarget);
            }
        }
        else
        {
            errorExcel = excelCount * 2 + 2;
            errorExcelLog = excelName + "#【实际模板（上一期）】#[" + isInertRowValue + "]未找到(序号出错)";
            errorList.Add((errorExcel, errorExcelLog, excelName));
            return errorList;
        }
        //数据修改
        var colCounMult = ExcelFixKeyDictionary[excelCount].Count;
        int colThreadCount = 8; // 线程数
        int colBatchSize = colCounMult / colThreadCount; // 每个线程处理的数据量
        Parallel.For(0, colThreadCount, e =>
        {
            int startIndex = e * colBatchSize;
            int endIndex = (e + 1) * colBatchSize;
            if (e == colThreadCount - 1) endIndex = colCounMult; // 最后一个线程处理剩余的数据
            for (int k = startIndex; k < endIndex; k++)
            {
                //查找字段所在列
                string excelKey = ExcelFixKeyDictionary[excelCount][k];
                int excelFileFixKey = FindSourceCol(sheet, 2, excelKey);
                //修改字段字典中的字段值，各自方法不一
                int rowThreadCount = 4; // 线程数
                int rowBatchSize = count / rowThreadCount; // 每个线程处理的数据量
                // 并发执行任务
                Parallel.For(0, rowThreadCount, i =>
                {
                    int startIndex = i * rowBatchSize;
                    int endIndex = (i + 1) * rowBatchSize;
                    if (i == rowThreadCount - 1) endIndex = count; // 最后一个线程处理剩余的数据

                    for (int j = startIndex; j < endIndex; j++)
                    {
                        //字典会把空值当0用
                        if (excelFileFixKey != -1)
                        {
                            var cellSource = sheet.Cells[endRowSource , excelFileFixKey];
                            var cellFix = sheet.Cells[endRowSource + j + 1, excelFileFixKey];
                            if (cellSource.Value == null)
                            {
                                continue;
                            }

                            if (cellSource.Value.ToString() == "")
                            {
                                continue;
                            }

                            // 字段每个数字位数统计，原始modeID统计
                            //KeyBitCount(cellFix.Value.ToString());
                            // 字段值改写方法
                            var temp1 = CellFixValueKeyList(ExcelFixKeyMethodDictionary[excelCount][k]);
                            // 修改字符串
                            var cellFixValue = StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                            if (cellFixValue == "^error^")
                            {
                                errorExcel = excelCount * 2 + 2;
                                errorExcel = excelCount * 2 + 2;
                                errorExcelLog = excelName + "#【修改模式】#[" + excelKey + "]字段方法写错";
                                errorList.Add((errorExcel, errorExcelLog, excelName));
                                return; // 终止当前线程
                            }

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
                    }
                });
            }

        });
        excel.Save();
        excel.Dispose();
        errorList.Add((errorExcel, errorExcelLog, excelName));
        return errorList;
    }

    public static string ExcelPathIgnore(dynamic excelPath, dynamic excelName)
    {
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

        return path;
    }
}

 public class ExcelDataInsertLanguage
 {
     public static void AutoInsertData()
     {
         dynamic App = ExcelDnaUtil.Application;
         var workBook = App.ActiveWorkbook;
         var excelPath = workBook.Path;
         var sourceSheet = workBook.Worksheets["多语言对话【模板】"];
         var fixSheet = workBook.Worksheets["数据修改"];
         var classSheet = workBook.Worksheets["枚举数据"];

         ErrorLogCtp.DisposeCtp();

         var errorExcelList = new List<List<(int, string, string)>>();

         List<(int, string, string)> error = null;

         error = LanguageDialogData(sourceSheet, fixSheet, classSheet, excelPath);

         if (error.Count != 0)
         {
             errorExcelList.Add(error);
         }

         //出错表格处理
         string errorLog = ExcelDataAutoInsert.ErrorExcelMark(errorExcelList, fixSheet);
         if (errorLog != "")
         {
             ErrorLogCtp.DisposeCtp();
             ErrorLogCtp.CreateCtpNormal(errorLog);
         }

         else
         {
             fixSheet.Range["A2:A1000"].Value = "";
         }
     }

     public static List<(int, string, string)> LanguageDialogData(dynamic sourceSheet,dynamic fixSheet,dynamic classSheet,string excelPath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var sourceData =PubMetToExcel.ExcelDataToList(sourceSheet);
        var sourcelTitle = sourceData.Item1;
        var sourceDataList = sourceData.Item2;

        var fixData = PubMetToExcel.ExcelDataToList(fixSheet);
        var fixlTitle = fixData.Item1;
        var fixDataList = fixData.Item2;

        var classData =PubMetToExcel.ExcelDataToList(classSheet);
        var classlTitle = classData.Item1;
        var classDataList = classData.Item2;

        var fileIndex = fixlTitle.IndexOf("表名");
        var keyIndex = fixlTitle.IndexOf("字段");
        var modelIdIndex = fixlTitle.IndexOf("初始模板");
        var newIdIndex = fixlTitle.IndexOf("实际模板(上一期)");

        int errorExcel =0;
        string errorExcelLog;
        var errorList = new List<(int, string, string)>();

        for (int i = 0; i < fixDataList.Count; i++)
        {
            if (fixDataList[i][fileIndex] == null) continue;
            //整理要修改的表格和原表格数据字段映射
            var sourceKeyList =new List<string>();
            var fixKeyList =new List<string>();
            for (int j = keyIndex; j < fixlTitle.Count; j++)
            {
                if (fixDataList[i][j] != null)
                {
                    var sourceKey = fixDataList[i][j].ToString();
                    sourceKeyList.Add(sourceKey);
                }
                if (fixDataList[i+1][j] != null)
                {
                    var fixKey = fixDataList[i+1][j].ToString();
                    fixKeyList.Add(fixKey);
                }
            }
            //遍历要修改的表格写入数据
            var fixFileName = fixDataList[i][fileIndex].ToString();
            var fixFileModeId = fixDataList[i][modelIdIndex].ToString();
            var fixFileNewId = fixDataList[i][newIdIndex].ToString();

            string path =ExcelDataAutoInsert.ExcelPathIgnore(excelPath, fixFileName);
            var targetExcel = new ExcelPackage(new FileInfo(path));
            ExcelWorkbook targetBook = null;
            try
            {
                targetBook = targetExcel.Workbook;
            }
            catch (Exception ex)
            {
                errorExcel = i * 2 + 2;
                errorExcelLog = fixFileName + "#不能创建WorkBook对象" + ex.Message;
                errorList.Add((errorExcel, errorExcelLog, fixFileName));
                continue;
            }

            ExcelWorksheet targetSheet = null;
            try
            {
                targetSheet = targetBook.Worksheets["Sheet1"] ?? targetBook.Worksheets[0];
            }
            catch (Exception ex)
            {
                errorExcel = i * 2 + 2;
                errorExcelLog = fixFileName + "#不能创建WorkBook对象" + ex.Message;
                errorList.Add((errorExcel, errorExcelLog, fixFileName));
                continue;
            }
            //数据复制
            var isInertRowTarget = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, fixFileNewId);
            //根据模板插入对应数据行，并复制
            var endRowSource = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, fixFileModeId);
            if (fixFileNewId != "")
            {
                if (isInertRowTarget == -1)
                {

                    targetSheet.InsertRow(endRowSource + 1, sourceDataList.Count);

                    var colCount = targetSheet.Dimension.Columns;
                    var cellSource = targetSheet.Cells[endRowSource, 1, endRowSource, colCount];
                    for (int m = 0; m < sourceDataList.Count; m++)
                    {
                        var cellTarget = targetSheet.Cells[endRowSource + 1 + m, 1, endRowSource + 1 + m, colCount];
                        cellSource.Copy(cellTarget, ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting | ExcelRangeCopyOptionFlags.ExcludeMergedCells);
                        cellSource.CopyStyles(cellTarget);
                    }
                }
            }

            //修改数据
            for (int m = 0; m < sourceDataList.Count; m++)
            {
                int sourceCount = 0;
                foreach (var source in sourceKeyList)
                {
                    var newstr = "";
                    if (int.TryParse(source, out var e))
                    {
                        var sourceValue = sourceDataList[m][sourcelTitle.IndexOf("GroupID")];
                        var str = sourceValue.ToString();
                        var digit = Math.Pow(10,e);
       
                        for (int k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourcelTitle.IndexOf("GroupID")];
                            if (repeatValue == sourceValue)
                            {
                                var newnum =long.Parse(str)* digit+k+1;
                                newstr = newstr + newnum+",";
                            }
                        }
                        newstr="["+newstr.Substring(0, newstr.Length - 1)+"]";
                        var cellCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, fixKeyList[sourceCount]);
                        if (cellCol == -1)
                        {
                            errorExcel = i * 2 + 2;
                            errorExcelLog = fixFileName + "#表格字段#[" + fixKeyList[sourceCount] + "]未找到";
                            errorList.Add((errorExcel, errorExcelLog, fixFileName));
                            continue;
                        }
                        var cellTarget = targetSheet.Cells[endRowSource + 1 + m, cellCol];
                        cellTarget.Value = newstr;
                    }
                    else if (source == "枚举1")
                    {
                        var sourceValue = sourceDataList[m][sourcelTitle.IndexOf("说话角色")];
                        var scCol = classlTitle.IndexOf(source);
                        var newID = "";
                        for (int k = 0; k< classDataList.Count; k++)
                        {
                            var targetVakye =classDataList[k][scCol];
                            if (targetVakye == sourceValue)
                            {
                                newID = classDataList[k][scCol+1].ToString();
                                break;
                            }
                        }
                        var cellCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, fixKeyList[sourceCount]);
                        if (cellCol == -1)
                        {
                            errorExcel = i * 2 + 2;
                            errorExcelLog = fixFileName + "#表格字段#[" + fixKeyList[sourceCount] + "]未找到";
                            errorList.Add((errorExcel, errorExcelLog, fixFileName));
                            continue;
                        }
                        var cellTarget = targetSheet.Cells[endRowSource + 1 + m, cellCol];
                        var sourceStr = cellTarget.Value.ToString();
                        var reg = "\\d+";
                        var matches = Regex.Matches(sourceStr, reg);
   
                        var oldID = matches[0].Value.ToString();
                        if (newID != "")
                        {
                            sourceStr = sourceStr.Replace(oldID, newID);
                        }
                        cellTarget.Value = sourceStr;
                    }
                    else
                    {
                        //GroupID不连续
                        var sourceValue = sourceDataList[m][sourcelTitle.IndexOf(source)]; 
                        var cellCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, fixKeyList[sourceCount]);
                        if (cellCol == -1)
                        {
                            errorExcel = i * 2 + 2;
                            errorExcelLog = fixFileName + "#表格字段#[" + fixKeyList[sourceCount] + "]未找到";
                            errorList.Add((errorExcel, errorExcelLog, fixFileName));
                            continue;
                        }
                        var cellTarget = targetSheet.Cells[endRowSource + 1 + m, cellCol];
                        cellTarget.Value = sourceValue;
                    }
                    sourceCount++;
                }
            }
            //数据去重
            if(errorExcel != 0) continue;
            if (fixFileName == "GuideDialogGroup.xlsx" )
            {
                int startRow = endRowSource+1;
                int endRow = startRow+ sourceDataList.Count;
                HashSet<string> uniqueValues = new HashSet<string>();
                // 遍历行并找到具有相同第一列值的行
                for (int row = startRow; row <= endRow; row++)
                {
                    string cellValue = targetSheet.Cells[row, 2].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        if (uniqueValues.Contains(cellValue))
                        {
                            // 如果发现第一列值相同的行，则删除该行
                            targetSheet.DeleteRow(row);
                            // 调整删除后的行号
                            row--;
                            endRow--;
                        }
                        else
                        {
                            uniqueValues.Add(cellValue);
                        }
                    }
                }
            }
            targetExcel.Save();
            targetExcel.Dispose();
        }
        return errorList;
    }
}