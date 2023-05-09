using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CommandBarButton = Microsoft.Office.Core.CommandBarButton;
using Color = System.Drawing.Color;
using MessageBox = System.Windows.Forms.MessageBox;
using LicenseContext = OfficeOpenXml.LicenseContext;
using Match = System.Text.RegularExpressions.Match;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Globalization;

namespace NumDesTools;

public class ExcelDataAutoInsert
{
    [ExcelFunction(IsHidden = true)]
    public static int FindTitle(dynamic sheet, int rows, string findValue)
    {
        var maxColumn = sheet.UsedRange.Columns.Count;
        for (var column = 1; column <= maxColumn; column++)
            if (sheet.Cells[rows, column] is Range cell && cell.Value2?.ToString() == findValue)
                return column;
        return -1;
    }

    public static int FindSourceCol(ExcelWorksheet sheet, int row, string searchValue)
    {
        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
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
        for (var row = 2; row <= sheet.Dimension.End.Row; row++)
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
    /*
        public static int FindKeyColNpoi(string excelPath, string targetWorkbook, int rows, string findValue, string targetSheet = "Sheet1")
        {
            string path;
            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
            switch (targetWorkbook)
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
                    path = excelPath + @"\" + targetWorkbook;
                    break;
            }
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
    */
    /*
        public static int FindKeyRowNpoi(string excelPath, string targetWorkbook, int cols, string findValue, string targetSheet = "Sheet1")
        {
            string path;
            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
            switch (targetWorkbook)
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
                    path = excelPath + @"\" + targetWorkbook;
                    break;
            }
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
    */
    /*
        private static List<(int, int)> ExcelDic(dynamic excelModeIdDictionary, dynamic excelModeIdNewDictionary, dynamic excelFixKeyDictionary, dynamic excelFixKeyMethodDictionary, dynamic excelFixGroup,dynamic sheet)
        {
            var modeCol = FindTitle(sheet, 1, "初始模板");
            var modeColNew = FindTitle(sheet, 1, "实际模板(上一期)");
            var excelCol = FindTitle(sheet, 1, "表名");
            var keyColFirst = FindTitle(sheet, 1, "修改字段");
            var addValueIndexMax = FindTitle(sheet, 1, "创建期号");
            var addValueIndexMin = FindTitle(sheet, 1, "模板期号");
            var addValue = Convert.ToInt32(sheet.Cells[2, addValueIndexMax].Value- sheet.Cells[2, addValueIndexMin].Value);
            var defaultData = new List<(int,int)> { (excelCol,addValue) };
            //读取模板表数据
            var rowsCount = sheet.UsedRange.Rows.Count;
            var colsCount = sheet.UsedRange.Columns.Count;
            var excelCount=0;
            for (var i = 2; i <= rowsCount; i++)
            {
                var cellExcel = sheet.Cells[i, excelCol].Value2;
                if (cellExcel == null) continue;
                var baseExcel = cellExcel.ToString();
                excelModeIdDictionary[excelCount] = new List<(string, string)>();
                excelModeIdNewDictionary[excelCount] =new List<(string, string)>();
                excelFixKeyDictionary[excelCount] = new List<string>();
                excelFixKeyMethodDictionary[excelCount] = new List<string>();
                excelFixGroup.Add(baseExcel);
                for (var j = keyColFirst; j <= colsCount; j++)
                {
                    //var baseExcelFixKeyCol = FindKeyColNPOI(excelPath, baseExcel, 1, baseExcelFixKey);
                    string baseExcelFixKey = sheet.Cells[i, j].Value2?.ToString() ?? "";
                    excelFixKeyDictionary[excelCount].Add(baseExcelFixKey);
                    var baseExcelFixKeyMethod = sheet.Cells[i + 1, j].Value2;
                    if (baseExcelFixKeyMethod == null)
                    {
                        baseExcelFixKeyMethod = "";
                    }
                    excelFixKeyMethodDictionary[excelCount].Add(baseExcelFixKeyMethod.ToString());
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
                excelModeIdDictionary[excelCount].Add(tuple);
                excelModeIdNewDictionary[excelCount].Add(tuple2);
                excelCount++;
            }
            return defaultData;
        }
    */
    /*
        private static List<(int, string, string)> SingleExcelDataWrite(int excelCount,int addValue,dynamic excelFixGroup, dynamic excelModeIdDictionary, dynamic excelModeIdNewDictionary, dynamic excelFixKeyDictionary, dynamic excelFixKeyMethodDictionary, dynamic excelPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excelName = excelFixGroup[excelCount];
            var startValue = excelModeIdDictionary[excelCount][0].Item1;
            var endValue = excelModeIdDictionary[excelCount][0].Item2;
            var isInertRowValue = excelModeIdNewDictionary[excelCount][0].Item1;
            var errorExcel=0;
            var errorExcelLog="";
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
            ExcelWorkbook workBook;
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
            ExcelWorksheet sheet;
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
            sheet ??= workBook.Worksheets[0];
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
            var countCc = 0;
            foreach (var keyIndex in excelFixKeyDictionary[excelCount])
            {
                if (keyIndex == "") continue;
                //查找字段所在列
                string excelKey = excelFixKeyDictionary[excelCount][countCc];
                var excelFileFixKey = FindSourceCol(sheet, 2, excelKey);
                //字典会把空值当0用
                if (excelFileFixKey == -1)
                {
                    countCc++;
                    continue;
                }
                //修改字段字典中的字段值，各自方法不一
                for (var i = 0; i < count; i++)
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
                    var temp1 = CellFixValueKeyList(excelFixKeyMethodDictionary[excelCount][countCc]);
                    //修改字符串
                    var cellFixValue = StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                    if (cellFixValue == "^error^")
                    {
                        errorExcel = excelCount * 2 + 2;
                        errorExcelLog = excelName + "#【修改模式】#[" + excelKey + "]字段方法写错";
                        errorList.Add((errorExcel, errorExcelLog, excelName));
                        return errorList;
                    }

                    cellFix.Value = double.TryParse(cellFixValue, out double number) ? number : cellFixValue;
                }
                countCc++;
            }
            excel.Save();
            excel.Dispose();
            errorList.Add((errorExcel, errorExcelLog, excelName));
            return errorList;
        }
    */

    /*
        public static string AutoInsertDat(bool threadMode)
        {
            dynamic app = ExcelDnaUtil.Application;
            var excelModeIdDictionary =new Dictionary<int, List<(string, string)>>();
            var excelModeIdNewDictionary = new Dictionary<int, List<(string, string)>>();
            var excelFixKeyDictionary =new Dictionary<int, List<string>>();
            var excelFixKeyMethodDictionary =new Dictionary<int, List<string>>();
            var excelFixGroup = new List<string>();
            var indexWk = app.ActiveWorkbook;
            var sheet = indexWk.ActiveSheet;
            var excelPath = indexWk.Path;

            ErrorLogCtp.DisposeCtp();

            var sw = new Stopwatch();
            sw.Start();
            //获取字典
            var defaultData=ExcelDic(excelModeIdDictionary,excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelFixGroup,sheet);
            var ts1 = Math.Round(sw.Elapsed.TotalSeconds,2);
            var str1 = "字典用时:" + ts1;
            //遍历文件
            var excelCount = 0;
             var errorExcelList = new List<List<(int,string,string)>>();
            foreach (var excelName in excelFixGroup)
            {
                List<(int, string,string)> error;
                if (threadMode)
                {
                    error = MultiExcelDataWrite(excelCount, defaultData[0].Item2, excelFixGroup, excelModeIdDictionary, excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelPath);
                }
                else
                {
                    error = SingleExcelDataWrite(excelCount, defaultData[0].Item2, excelFixGroup, excelModeIdDictionary, excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelPath);
                }
                if (error.Count != 0)
                {
                    errorExcelList.Add(error);
                }
                excelCount++;
                app.StatusBar = "写入数据" + "<" + excelCount + "/" + excelFixGroup.Count + ">" + excelName;
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
    */
    [ExcelFunction(IsHidden = true)]
    public static string ErrorExcelMark(dynamic errorExcelList, dynamic sheet)
    {
        var strBuild = new StringBuilder();
        for (var i = 0; i < errorExcelList.Count; i++)
        {
            if (errorExcelList[i][0].Item1 == 0) continue;
            strBuild.Append(errorExcelList[i][0].Item2);
            var cell = sheet.Cells[errorExcelList[i][0].Item1, 1];
            cell.Value = "git checkout -- Excels/Tables/" + errorExcelList[i][0].Item3;
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

    public static string StringRegPlace(string str, List<(int, int)> digit, int addValue)
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
            var index = match2.Index;
            var num = long.Parse(numStr);
            if (digit.Any(item => item.Item1 == matchCount + 1))
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
                var addDigit = Math.Abs((long)Math.Pow(10, digit[0].Item2 - 1) * addValue);
                if (addDigit > num*100)
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
            var modeCol = FindTitle(sheet, 1, "实际模板(上一期)");
            var excelName = FindTitle(sheet, 1, "表名");
            string findValue = sheet.Cells[i, modeCol].Value?.ToString();
            var cell = sheet.Cells[i, excelName];
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx")) continue;
            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
            string path = cell.value switch
            {
                "Localizations.xlsx" => newPath + @"\Excels\Localizations\Localizations.xlsx",
                "UIConfigs.xlsx" => newPath + @"\Excels\UIs\UIConfigs.xlsx",
                "UIItemConfigs.xlsx" => newPath + @"\Excels\UIs\UIItemConfigs.xlsx",
                _ => excelPath + @"\" + cell.value
            };

            var excel = new ExcelPackage(new FileInfo(path));
            var workbook = excel.Workbook;
            var sheetTemp = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];
            var row = FindSourceRow(sheetTemp, 2, findValue);
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
    public static void ExcelHyperLinksNormal(dynamic excelPath, dynamic sheet)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        for (var i = 2; i <= 500; i++)
        {
            var cell = sheet.Cells[i, 5];
            if (cell.value == null || !cell.value.ToString().Contains(".xlsx")) continue;
            var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
            string path = cell.value switch
            {
                "Localizations.xlsx" => newPath + @"\Excels\Localizations\Localizations.xlsx",
                "UIConfigs.xlsx" => newPath + @"\Excels\UIs\UIConfigs.xlsx",
                "UIItemConfigs.xlsx" => newPath + @"\Excels\UIs\UIItemConfigs.xlsx",
                _ => excelPath + @"\" + cell.value
            };

            //var excel = new ExcelPackage(new FileInfo(path));
            //var workbook = excel.Workbook;
            //var sheetTemp = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

            //var sheetName = sheetTemp.Name;
                var links = path + "#" + "Sheet1!A1";
                //excel.Dispose();
                cell.Hyperlinks.Add(cell, links);
                cell.Font.Size = 9;
                cell.Font.Name = "微软雅黑";
            
        }
    }
    /*
        public static void CellFormatAuto(dynamic excelModeIdDictionary, dynamic excelModeIdNewDictionary, dynamic excelFixKeyDictionary, dynamic excelFixKeyMethodDictionary, dynamic excelFixGroup, dynamic sheet)
        {
            var defaultData = ExcelDic(excelModeIdDictionary, excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelFixGroup, sheet);
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
    */
    /*
        public static void RightClickWriteExcel(CommandBarButton ctrl, ref bool cancelDefault)
        {
            dynamic app = ExcelDnaUtil.Application;
            var excelModeIdDictionary = new Dictionary<int, List<(string, string)>>();
            var excelModeIdNewDictionary = new Dictionary<int, List<(string, string)>>();
            var excelFixKeyDictionary = new Dictionary<int, List<string>>();
            var excelFixKeyMethodDictionary = new Dictionary<int, List<string>>();
            var excelFixGroup = new List<string>();
            var indexWk = app.ActiveWorkbook;
            var sheet = indexWk.ActiveSheet;
            var excelPath = indexWk.Path;

            ErrorLogCtp.DisposeCtp();

            var sw = new Stopwatch();
            sw.Start();
            var defaultData = ExcelDic(excelModeIdDictionary, excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelFixGroup, sheet);
            var errorExcelList = new List<List<(int, string,string)>>();

            var cell = app.Selection;
            var rowStart = cell.Row;
            var rowCount =cell.Rows.Count;
            var rowEnd = rowStart + rowCount - 1;
            for (int i = rowStart; i <= rowEnd; i++)
            {
                var realExcel =(i-2) % 2;
                if (realExcel == 0)
                {
                    var excelCount = (i - 2) / 2;
                    var error = SingleExcelDataWrite(excelCount, defaultData[0].Item2, excelFixGroup, excelModeIdDictionary,
                        excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelPath);
                    errorExcelList.Add(error);
                    app.StatusBar = "写入数据" + "<" + excelCount + "/" + excelFixGroup.Count + ">" + excelFixGroup[excelCount];
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
            app.StatusBar = str2;
        }
    */
    /*
        public static void RightClickWriteExcelThread(CommandBarButton ctrl, ref bool cancelDefault)
        {
            dynamic app = ExcelDnaUtil.Application;
            var excelModeIdDictionary = new Dictionary<int, List<(string, string)>>();
            var excelModeIdNewDictionary = new Dictionary<int, List<(string, string)>>();
            var excelFixKeyDictionary = new Dictionary<int, List<string>>();
            var excelFixKeyMethodDictionary = new Dictionary<int, List<string>>();
            var excelFixGroup = new List<string>();
            var indexWk = app.ActiveWorkbook;
            var sheet = indexWk.ActiveSheet;
            var excelPath = indexWk.Path;

            ErrorLogCtp.DisposeCtp();

            var sw = new Stopwatch();
            sw.Start();
            var defaultData = ExcelDic(excelModeIdDictionary, excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelFixGroup, sheet);
            var errorExcelList = new List<List<(int, string,string)>>();

            var cell = app.Selection;
            var rowStart = cell.Row;
            var rowCount = cell.Rows.Count;
            var rowEnd = rowStart + rowCount - 1;
            for (int i = rowStart; i <= rowEnd; i++)
            {
                var realExcel = (i - 2) % 2;
                if (realExcel == 0)
                {
                    var excelCount = (i - 2) / 2;
                    var error = MultiExcelDataWrite(excelCount, defaultData[0].Item2, excelFixGroup, excelModeIdDictionary,
                        excelModeIdNewDictionary, excelFixKeyDictionary, excelFixKeyMethodDictionary, excelPath);

                    errorExcelList.Add(error);
                    app.StatusBar = "写入数据" + "<" + excelCount + "/" + excelFixGroup.Count + ">" + excelFixGroup[excelCount];
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
            app.StatusBar = str2;
        }
    */
    /*
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
    */

    public static List<(int, int)> CellFixValueKeyList(string str)
    {
        var monkeyList = new List<(int, int)>();

        str ??= "";

        if (str.Contains(','))
        {
            var pairs = str.Split(',');
            foreach (var pair in pairs)
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

    /*
        private static List<(int, string, string)> MultiExcelDataWrite(int excelCount, int addValue, dynamic excelFixGroup, dynamic excelModeIdDictionary, dynamic excelModeIdDNewictionary, dynamic excelFixKeyDictionary, dynamic excelFixKeyMethodDictionary, dynamic excelPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excelName = excelFixGroup[excelCount];
            var startValue = excelModeIdDictionary[excelCount][0].Item1;
            var endValue = excelModeIdDictionary[excelCount][0].Item2;
            var isInertRowValue = excelModeIdDNewictionary[excelCount][0].Item1;
            var errorExcel =0;
            var errorExcelLog="";
            var errorList = new List<(int, string,string)>();
            string path = ExcelPathIgnore(excelPath, excelName);
            var excel = new ExcelPackage(new FileInfo(path));
            ExcelWorkbook workBook;
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

            ExcelWorksheet sheet;
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
            var colCounMult = excelFixKeyDictionary[excelCount].Count;
            var colThreadCount = 8; // 线程数
            int colBatchSize = colCounMult / colThreadCount; // 每个线程处理的数据量
            Parallel.For(0, colThreadCount, e =>
            {
                var startRow = e * colBatchSize;
                var endRow = (e + 1) * colBatchSize;
                if (e == colThreadCount - 1) endRow = colCounMult; // 最后一个线程处理剩余的数据
                for (var k = startRow; k < endRow; k++)
                {
                    //查找字段所在列
                    string excelKey = excelFixKeyDictionary[excelCount][k];
                    var excelFileFixKey = FindSourceCol(sheet, 2, excelKey);
                    //修改字段字典中的字段值，各自方法不一
                    var rowThreadCount = 4; // 线程数
                    int rowBatchSize = count / rowThreadCount; // 每个线程处理的数据量
                    // 并发执行任务
                    var k1 = k;
                    Parallel.For(0, rowThreadCount, i =>
                    {
                        var startCol = i * rowBatchSize;
                        var endCol= (i + 1) * rowBatchSize;
                        if (i == rowThreadCount - 1) endCol = count; // 最后一个线程处理剩余的数据

                        for (var j = startCol; j < endCol; j++)
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
                                var temp1 = CellFixValueKeyList(excelFixKeyMethodDictionary[excelCount][k1]);
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

                                if (double.TryParse(cellFixValue, out double number))
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
    */
    [ExcelFunction(IsHidden = true)]
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
        dynamic app = ExcelDnaUtil.Application;
        var workBook = app.ActiveWorkbook;
        var excelPath = workBook.Path;
        var sourceSheet = workBook.Worksheets["多语言对话【模板】"];
        var fixSheet = workBook.Worksheets["数据修改"];
        var classSheet = workBook.Worksheets["枚举数据"];
        var emoSheet = workBook.Worksheets["表情枚举"];

        ErrorLogCtp.DisposeCtp();

        var errorExcelList = new List<List<(int, string, string)>>();

        List<(int, string, string)> error = LanguageDialogData(sourceSheet, fixSheet, classSheet,emoSheet, excelPath, app);

        if (error.Count != 0) errorExcelList.Add(error);

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

    public static List<(int, string, string)> LanguageDialogData(dynamic sourceSheet, dynamic fixSheet,
        dynamic classSheet, dynamic emoSheet,string excelPath, dynamic app)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var sourceData = PubMetToExcel.ExcelDataToList(sourceSheet);
        var sourceTitle = sourceData.Item1;
        var sourceDataList = sourceData.Item2;

        var fixData = PubMetToExcel.ExcelDataToList(fixSheet);
        var fixTitle = fixData.Item1;
        var fixDataList = fixData.Item2;

        var classData = PubMetToExcel.ExcelDataToList(classSheet);
        var classTitle = classData.Item1;
        var classDataList = classData.Item2;

        var emoData = PubMetToExcel.ExcelDataToList(emoSheet);
        var emoDataList = emoData.Item2;

        var fileIndex = fixTitle.IndexOf("表名");
        var keyIndex = fixTitle.IndexOf("字段");
        var modelIdIndex = fixTitle.IndexOf("初始模板");

        var errorExcel = 0;
        var errorList = new List<(int, string, string)>();

        for (var i = 0; i < fixDataList.Count; i++)
        {
            if (fixDataList[i][fileIndex] == null) continue;
            //整理要修改的表格和原表格数据字段映射
            var sourceKeyList = new List<string>();
            var fixKeyList = new List<string>();
            for (int j = keyIndex; j < fixTitle.Count; j++)
            {
                if (fixDataList[i][j] != null)
                {
                    var sourceKey = fixDataList[i][j].ToString();
                    sourceKeyList.Add(sourceKey);
                }

                if (fixDataList[i + 1][j] != null)
                {
                    var fixKey = fixDataList[i + 1][j].ToString();
                    fixKeyList.Add(fixKey);
                }
            }

            //遍历要修改的表格写入数据
            var fixFileName = fixDataList[i][fileIndex].ToString();
            var fixFileModeId = fixDataList[i][modelIdIndex].ToString();

            string path = ExcelDataAutoInsert.ExcelPathIgnore(excelPath, fixFileName);
            var targetExcel = new ExcelPackage(new FileInfo(path));
            ExcelWorkbook targetBook;
            string errorExcelLog;
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

            ExcelWorksheet targetSheet;
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
            //数据查重
            var c = 0;
            if (fixFileName == "GuideDialogDetail.xlsx")
            {
                c = 1;
            }
            else if (fixFileName == "Localizations.xlsx")
            {
                c = 2;
            }
            else if (fixFileName == "GuideDialogBranch.xlsx")
            {
                c = 3;
            }
            var idList = new List<string>();
            for (int r = 0; r < sourceDataList.Count; r++)
            {
                var value = sourceDataList[r][c]?.ToString() ?? "";
                idList.Add(value);
            }
            var newIdList = idList.Distinct().ToList();

            // 定义要删除的行的列表
            List<int> rowsToDelete = new List<int>();
            foreach (var id in newIdList)
            {
                var reDd = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, id);
                if (reDd != -1)
                {
                    rowsToDelete.Add(reDd);
                }
            }

            //int endRow = targetSheet.Dimension.End.Row;
            //// 遍历行并找到具有相同第一列值的行
            //for (var row = 4; row <= endRow; row++)
            //{
            //    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";
            //    if (idSet.Contains(cellValue))
            //    {
            //        // 如果发现第一列值相同的行，则删除该行
            //        targetSheet.DeleteRow(row);
            //        // 调整删除后的行号
            //        row--;
            //        endRow--;
            //    }
            //}
            rowsToDelete.Sort();
            rowsToDelete.Reverse();
            // 逐行删除
            foreach (var rowToDelete in rowsToDelete)
            {
                targetSheet.DeleteRow(rowToDelete, 1);
            }

            //根据模板插入对应数据行，并复制
            var endRowSource = ExcelDataAutoInsert.FindSourceRow(targetSheet, 2, fixFileModeId);
            if (endRowSource == -1)
            {
                MessageBox.Show(fixFileModeId+@"目标表中不存在");
                continue;
            }
            targetSheet.InsertRow(endRowSource + 1, sourceDataList.Count);
            var colCount = targetSheet.Dimension.Columns;
            var cellSource = targetSheet.Cells[endRowSource, 1, endRowSource, colCount];
            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var cellTarget = targetSheet.Cells[endRowSource + 1 + m, 1, endRowSource + 1 + m, colCount];
                cellSource.Copy(cellTarget,
                    ExcelRangeCopyOptionFlags.ExcludeConditionalFormatting |
                    ExcelRangeCopyOptionFlags.ExcludeMergedCells);
                cellSource.CopyStyles(cellTarget);
            }

            //修改数据
            for (var m = 0; m < sourceDataList.Count; m++)
            {
                var sourceCount = 0;
                foreach (var source in sourceKeyList)
                {
                    var cellCol = ExcelDataAutoInsert.FindSourceCol(targetSheet, 2, fixKeyList[sourceCount]);
                    if (cellCol == -1)
                    {
                        if (fixKeyList[sourceCount] == "bgType")
                        {
                            sourceCount++;
                            continue;
                        }
                        
                        errorExcel = i * 2 + 2;
                        errorExcelLog = fixFileName + "#表格字段#[" + fixKeyList[sourceCount] + "]未找到";
                        errorList.Add((errorExcel, errorExcelLog, fixFileName));
                        sourceCount++;
                        continue;
                    }

                    var cellTarget = targetSheet.Cells[endRowSource + 1 + m, cellCol];
                    var newStr = "";
                    if (int.TryParse(source, out var e))
                    {
                        var realCol = "";
                        if (fixFileName == "GuideDialogGroup.xlsx")
                            realCol = "GroupID";
                        else if (fixFileName == "GuideDialogBranch.xlsx") realCol = "BranchID";
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(realCol)];
                        if (sourceValue == "" || sourceValue == null) continue;
                        var str = sourceValue.ToString();
                        var digit = Math.Pow(10, e);
                        var repeatCount = 0;
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf(realCol)];
                            if (repeatValue == "" || repeatValue == null) continue;
                            if (repeatValue == sourceValue)
                            {
                                var newNum = long.Parse(str) * digit + repeatCount + 1;
                                newStr = newStr + newNum + ",";
                                repeatCount++;
                            }
                        }

                        newStr = "[" + newStr.Substring(0, newStr.Length - 1) + "]";
                        cellTarget.Value = newStr;
                    }
                    else if (source == "枚举1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var scCol = classTitle.IndexOf(source);
                        var newId = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                newId = classDataList[k][scCol + 1].ToString();
                                break;
                            }
                        }

                        var sourceStr = cellTarget.Value?.ToString();
                        var reg = "\\d+";
                        if (sourceStr == null || sourceStr == "") continue;
                        var matches = Regex.Matches(sourceStr, reg);

                        var oldId = matches[0].Value.ToString();
                        if (newId != "") sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source =="角色表情")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)];
                        for (var k = 0; k < emoDataList.Count; k++)
                        {
                            var targetValue = emoDataList[k][0];
                            if (targetValue == sourceValue)
                            {
                                var emoId = emoDataList[k][2];
                                cellTarget.Value = emoId;
                                break;
                            }
                        }
                    }
                    else if (source == "触发分支")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)]?.ToString();
                        if (sourceValue == null || sourceValue == "" || sourceValue == "0")
                        {
                            sourceCount++;
                            continue;
                        }
                        var uniqueValues1 = new HashSet<string>();
                        var strBranch = "";
                        for (var k = 0; k < sourceDataList.Count; k++)
                        {
                            var repeatValue = sourceDataList[k][sourceTitle.IndexOf("分支归属")]?.ToString();
                            if (repeatValue == null || repeatValue == "") continue;
                            if (repeatValue == sourceValue)
                            {
                                var branchId = sourceDataList[k][sourceTitle.IndexOf("BranchID")];
                                if (!uniqueValues1.Contains(branchId))
                                {
                                    uniqueValues1.Add(branchId);
                                    strBranch = strBranch + branchId + ",";
                                }
                            }
                        }

                        strBranch = "[" + strBranch.Substring(0, strBranch.Length - 1) + "]";
                        cellTarget.Value = strBranch;
                    }
                    else if (source == "分支多语言")
                    {
                        var newId = sourceDataList[m][sourceTitle.IndexOf("BranchID")]?.ToString();
                        var sourceStr = cellTarget.Value?.ToString();
                        if (sourceStr == null || sourceStr == "") continue;
                        var reg = "\\d+";
                        var matches = Regex.Matches(sourceStr, reg);
                        var oldId = matches[0].Value.ToString();
                        if (newId != "") sourceStr = sourceStr.Replace(oldId, newId);
                        cellTarget.Value = sourceStr;
                    }
                    else if (source == "角色换装1")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m][sourceTitle.IndexOf("角色换装")]?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                if (sourceValue2 == "1")
                                {
                                    newValue = classDataList[k][scCol + 2].ToString();
                                }
                                else
                                {
                                    newValue = "[]";
                                }
                                break;
                            }
                        }
                        cellTarget.Value = newValue;
                    }
                    else if (source == "角色换装2")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("说话角色")];
                        var sourceValue2 = sourceDataList[m][sourceTitle.IndexOf("角色换装")]?.ToString();
                        var scCol = classTitle.IndexOf("枚举1");
                        var newValue = "";
                        for (var k = 0; k < classDataList.Count; k++)
                        {
                            var targetValueKey = classDataList[k][scCol];
                            if (targetValueKey == sourceValue)
                            {
                                if (sourceValue2 != "1")
                                {
                                    newValue = classDataList[k][scCol + 3].ToString();
                                }
                                else
                                {
                                    newValue = "";
                                }
                                break;
                            }
                        }
                        cellTarget.Value = newValue;
                    }
                    else if (source == "UI对话框")
                    {
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf("UI对话框")];
                        if (sourceValue == null)
                        {
                            sourceValue = "1";
                        }
                        else
                        {
                            sourceValue = sourceValue.ToString();
                        }
                        if (fixKeyList[sourceCount] == "bgType")
                        {
                            cellTarget.Value = sourceValue;
                        }
                    }
                    else
                    {
                        //GroupID不连续
                        var sourceValue = sourceDataList[m][sourceTitle.IndexOf(source)];
                        cellTarget.Value = sourceValue;
                    }

                    sourceCount++;
                }
            }

            //数据合并
            if (errorExcel != 0) continue;
            int startRow = endRowSource + 1;
            int endRow2 = startRow + sourceDataList.Count - 1;
            //int endRow2 = targetSheet.Dimension.End.Row;
            //if (hasCopy == 2)
            //{
            //    var dataCount = int.Parse(fixDataList[i][dataRows]);
            //    var start = endRow2 + 1;
            //    targetSheet.DeleteRow(start, dataCount);
            //}
            if (fixFileName == "GuideDialogBranch.xlsx" || fixFileName == "GuideDialogGroup.xlsx")
            {
                var uniqueValues = new HashSet<string>();
                // 遍历行并找到具有相同第一列值的行
                for (var row = 4; row <= endRow2; row++)
                {
                    var cellValue = targetSheet.Cells[row, 2].Value?.ToString() ?? "";

                    if (uniqueValues.Contains(cellValue) || cellValue == "")
                    {
                        // 如果发现第一列值相同的行，则删除该行
                        targetSheet.DeleteRow(row);
                        // 调整删除后的行号
                        row--;
                        endRow2--;
                    }
                    else
                    {
                        uniqueValues.Add(cellValue);
                    }
                }
            }
            targetExcel.Save();
            targetExcel.Dispose();
            var excelCount = i/2 + 1;
            app.StatusBar = "写入数据" + "<" + excelCount + "/" + fixDataList.Count / 2 + ">" + fixFileName;
        }

        return errorList;
    }
}

public class ExcelDataAutoInsertMulti
{
    public static void InsertData(dynamic isMulti)
    {
        dynamic app = ExcelDnaUtil.Application;
        var indexWk = app.ActiveWorkbook;
        var sheet = app.ActiveSheet;
        var excelPath = indexWk.Path;
        var colsCount = sheet.UsedRange.Columns.Count;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdCol = title.IndexOf("初始模板");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        var fixKeyCol = title.IndexOf("修改字段");
        var baseIdCol = title.IndexOf("模板期号");
        var creatIdCol = title.IndexOf("创建期号");
        var commentValue = data[2][baseIdCol];
        //var cellBackColor = data[4][baseIdCol];
        var writeMode = data[2][creatIdCol];
        ErrorLogCtp.DisposeCtp();
        //获取单元格颜色
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        //ID自增跨度
        var addValue = (int)data[0][creatIdCol] - (int)data[0][baseIdCol];
        //字典Value跨度（行）
        var rowCount = 2;
        //获取字典
        var colFixKeyCount = colsCount - fixKeyCol;
        var modelId = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdCol, rowCount);
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdNewCol, rowCount);
        var fixKey = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, fixKeyCol, rowCount, colFixKeyCount);
        var ignoreExcel = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, creatIdCol, rowCount);
        //遍历文件写入
        var errorExcelList = new List<List<(string, string, string)>>();
        var excelCount = 1;
        foreach (var key in modelId)
        {
            //写入算法
            var excelName = key.Key;
            //过滤不导出的表格
            var ignore = ignoreExcel[excelName][0].Item1[0, 0];
            if (ignore != null)
            {
                var ignoreStr = ignore.ToString();
                if (ignoreStr == "跳过")
                {
                    app.StatusBar = "跳过" + "<" + excelName;
                    excelCount++;
                    continue;
                }
            }
            List<(string, string, string)> error =
                ExcelDataWrite(modelId, modelIdNew, fixKey, excelPath, excelName, addValue, isMulti, commentValue,
                    cellColor, writeMode);
            app.StatusBar = "写入数据" + "<" + excelCount + "/" + modelId.Count + ">" + excelName;
            errorExcelList.Add(error);
            excelCount++;
        }

        //错误日志处理
        var errorLog = ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //sheet.Range["B4"].Value = "否";
            app.StatusBar = "完成写入";
            return;
        }

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
    }

    public static void RightClickInsertData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sw = new Stopwatch();
        sw.Start();

        dynamic app = ExcelDnaUtil.Application;
        var indexWk = app.ActiveWorkbook;
        var sheet = app.ActiveSheet;
        var excelPath = indexWk.Path;
        var colsCount = sheet.UsedRange.Columns.Count;
        var sheetData = PubMetToExcel.ExcelDataToList(sheet);
        var title = sheetData.Item1;
        var data = sheetData.Item2;
        var sheetNameCol = title.IndexOf("表名");
        var modelIdCol = title.IndexOf("初始模板");
        var modelIdNewCol = title.IndexOf("实际模板(上一期)");
        var fixKeyCol = title.IndexOf("修改字段");
        var baseIdCol = title.IndexOf("模板期号");
        var creatIdCol = title.IndexOf("创建期号");
        var commentValue = data[2][baseIdCol];
        //var cellBackColor = data[4][baseIdCol];
        var writeMode = data[2][creatIdCol];
        ErrorLogCtp.DisposeCtp();
        //获取单元格颜色
        var colorCell = sheet.Cells[6, 1];
        var cellColor = PubMetToExcel.GetCellBackgroundColor(colorCell);
        //ID自增跨度
        var addValue = (int)data[0][creatIdCol] - (int)data[0][baseIdCol];
        //字典Value跨度（行）
        var rowCount = 2;
        //获取字典
        var colFixKeyCount = colsCount - fixKeyCol;
        var modelId = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdCol, rowCount);
        var modelIdNew = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, modelIdNewCol, rowCount);
        var fixKey = PubMetToExcel.ExcelDataToDictionary(data, sheetNameCol, fixKeyCol, rowCount, colFixKeyCount);
        //遍历文件写入
        var errorExcelList = new List<List<(string, string, string)>>();
        var cell = app.Selection;
        var rowStart = cell.Row;
        var rowCountNew = cell.Rows.Count;
        var rowEnd = rowStart + rowCountNew - 1;
        var excelList = new List<string>();
        //获得一导出文件集合
        for (int i = rowStart; i <= rowEnd; i++)
        {
            var excelName = data[i - 2][sheetNameCol];
            excelList.Add(excelName);
        }

        //去重
        var newExcelList = excelList.Where(excelName => !string.IsNullOrEmpty(excelName)).Distinct().ToList();
        for (var i = 0; i < newExcelList.Count; i++)
        {
            //写入算法
            var excelName = newExcelList[i];
            if (excelName == null) continue;
            List<(string, string, string)> error =
                ExcelDataWrite(modelId, modelIdNew, fixKey, excelPath, excelName, addValue, false, commentValue,
                    cellColor, writeMode);
            app.StatusBar = "写入数据" + "<" + i + "/" + newExcelList.Count + ">" + excelName;
            errorExcelList.Add(error);
        }

        //错误日志处理
        var errorLog = ErrorLogAnalysis(errorExcelList, sheet);
        if (errorLog == "")
        {
            //sheet.Range["B4"].Value = "否";
            sw.Stop();
            var ts2 = Math.Round(sw.Elapsed.TotalSeconds, 2);
            app.StatusBar = "完成写入，用时：" + ts2.ToString(CultureInfo.InvariantCulture);
            return;
        }
        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);
        sw.Stop();
        var ts3 = Math.Round(sw.Elapsed.TotalSeconds, 2);
        app.StatusBar = "完成写入:有错误，用时：" + ts3.ToString(CultureInfo.InvariantCulture);
    }

    public static List<(string, string, string)> ExcelDataWrite(dynamic modelId, dynamic modelIdNew, dynamic fixKey,
        dynamic excelPath, dynamic excelName, dynamic addValue, dynamic modeThread, dynamic commentValue,
        dynamic cellBackColor, dynamic writeMode)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

        var excel = new ExcelPackage(new FileInfo(path));
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

        ExcelWorksheet sheet;
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
        //获取要查重的ID
        var writeIdList = ExcelDataWriteIdGroup(excelName, addValue, sheet, fixKey, modelId);
        //执行查重删除
        PubMetToExcel.RepeatValue2(sheet,4,2,writeIdList.Item1);
        var colCount = sheet.Dimension.Columns;
        //第一次写入插入行的位置，因为可能删除导致行数变化，需要重新获取一次
        writeIdList = ExcelDataWriteIdGroup(excelName, addValue, sheet, fixKey, modelId);
        var writeRow = writeIdList.Item2;
        //执行写入操作
        for (var excelMulti = 0; excelMulti < modelId[excelName].Count; excelMulti++)
        {
            var startValue = modelId[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelId[excelName][excelMulti].Item1[1, 0].ToString();
            //var writeRow = sheet.Dimension.End.Row;

            var startRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, startValue);
            if (startRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + startValue + "]未找到(序号出错)";
                errorList.Add((startValue, errorExcelLog, excelName));
                return errorList;
            }

            var endRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, endValue);
            if (endRowSource == -1)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]未找到(序号出错)";
                errorList.Add((endValue, errorExcelLog, excelName));
                return errorList;
            }

            if (endRowSource - startRowSource < 0)
            {
                errorExcelLog = excelName + "#【初始模板】#[" + endValue + "]起始、终结ID顺序反了";
                errorList.Add((endValue, errorExcelLog, excelName));
                return errorList;
            }

            //数据复制
            var count = endRowSource - startRowSource + 1;
            sheet.InsertRow(writeRow + 1, count);
            var cellSource = sheet.Cells[startRowSource, 1, endRowSource, colCount];
            var cellTarget = sheet.Cells[writeRow + 1, 1, writeRow + count, colCount];
            cellTarget.Value = cellSource.Value;
            //设置背景色
            cellTarget.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cellTarget.Style.Fill.BackgroundColor.SetColor(cellBackColor);

            //数据修改
            var fixItem = fixKey[excelName][excelMulti].Item1;
            errorList = modeThread
                ? (List<(string, string, string)>)MultiWrite(excelName, addValue, fixItem, sheet, count, startRowSource, errorList, commentValue,writeRow)
                : (List<(string, string, string)>)SingleWrite(excelName, addValue, fixItem, sheet, count, startRowSource, errorList, commentValue, writeRow);
            writeRow+=count;
        }

        excel.Save();
        excel.Dispose();
        errorList.Add(("-1", errorExcelLog, excelName));
        return errorList;
    }

    private static List<(string, string, string)> SingleWrite(dynamic excelName, dynamic addValue, dynamic fixItem,
        ExcelWorksheet sheet,
        dynamic count, dynamic startRowSource, List<(string, string, string)> errorList,
        dynamic commentValue, int writeRow)
    {
        for (var colMulti = 0; colMulti < fixItem.GetLength(1); colMulti++)
        {
            string excelKey = fixItem[0, colMulti];
            if (excelKey == null) continue;
            var excelFileFixKey = ExcelDataAutoInsert.FindSourceCol(sheet, 2, excelKey);
            if (excelFileFixKey == -1)
            {
                var errorExcelLog = excelName + "#【初始模板】#[" + excelKey + "]未找到(字段出错)";
                errorList.Add((excelKey, errorExcelLog, excelName));
                continue;
            }

            string excelKeyMethod = fixItem[1, colMulti]?.ToString();
            //修改字段字典中的字段值，各自方法不一
            for (var i = 0; i < count; i++)
            {
                var cellSource = sheet.Cells[startRowSource + i, excelFileFixKey];
                var rowId = sheet.Cells[startRowSource + i, 2];
                var cellCol = sheet.Cells[2, excelFileFixKey].Value?.ToString();
                var cellFix = sheet.Cells[writeRow + 1 + i, excelFileFixKey];
                if (cellSource.Value == null) continue;

                if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0") continue;

                if (cellCol != null && cellCol.Contains("#"))
                {
                    cellFix.Value = commentValue;
                }
                else
                {
                    //字段值改写方法
                    var temp1 = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                    //修改字符串
                    var cellFixValue = ExcelDataAutoInsert.StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                    if (cellFixValue == "^error^")
                    {
                        string errorExcelLog = excelName +"#" +rowId.Value + "#【修改模式】#[" + excelKey + "]字段方法写错";
                        errorList.Add((excelKey, errorExcelLog, excelName));
                    }
                    cellFix.Value = double.TryParse(cellFixValue, out double number) ? number : cellFixValue;
                }
            }
        }

        return errorList;
    }

    private static List<(string, string, string)> MultiWrite(dynamic excelName, dynamic addValue, dynamic fixItem,
        ExcelWorksheet sheet,
        dynamic count, dynamic startRowSource, List<(string, string, string)> errorList,
        dynamic commentValue, int writeRow)
    {
        var colCoinMulti = fixItem.GetLength(1);
        var colThreadCount = 8; // 线程数
        int colBatchSize = colCoinMulti / colThreadCount; // 每个线程处理的数据量
        Parallel.For(0, colThreadCount, e =>
        {
            var startRow = e * colBatchSize;
            var endRow = (e + 1) * colBatchSize;
            if (e == colThreadCount - 1) endRow = colCoinMulti; // 最后一个线程处理剩余的数据
            for (var k = startRow; k < endRow; k++)
            {
                //查找字段所在列
                string excelKey = fixItem[0, k];
                if (excelKey == null) continue;
                var excelFileFixKey = ExcelDataAutoInsert.FindSourceCol(sheet, 2, excelKey);
                if (excelFileFixKey == -1)
                {
                    var errorExcelLog = excelName + "#【初始模板】#[" + excelKey + "]未找到(字段出错)";
                    errorList.Add((excelKey, errorExcelLog, excelName));
                    continue;
                }

                string excelKeyMethod = fixItem[1, k]?.ToString();

                var rowThreadCount = 4; // 线程数
                int rowBatchSize = count / rowThreadCount; // 每个线程处理的数据量
                // 并发执行任务
                Parallel.For(0, rowThreadCount, i =>
                {
                    var startCol = i * rowBatchSize;
                    var endCol = (i + 1) * rowBatchSize;
                    if (i == rowThreadCount - 1) endCol = count; // 最后一个线程处理剩余的数据

                    for (var j = startCol; j < endCol; j++)
                    {
                        var cellSource = sheet.Cells[startRowSource + j, excelFileFixKey];
                        var cellCol = sheet.Cells[2, excelFileFixKey].Value?.ToString();
                        var cellFix = sheet.Cells[writeRow + j + 1, excelFileFixKey];
                        var rowId = sheet.Cells[startRowSource + j, 2];
                        if (cellSource.Value == null) continue;

                        if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0") continue;

                        if (cellCol != null && cellCol.Contains("#"))
                        {
                            cellFix.Value = commentValue;
                        }
                        else
                        {
                            //字段值改写方法
                            var temp1 = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                            //修改字符串
                            var cellFixValue =
                                ExcelDataAutoInsert.StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                            if (cellFixValue == "^error^")
                            {
                                string errorExcelLog = excelName + "#" + rowId.Value + "#【修改模式】#[" + excelKey + "]字段方法写错";
                                errorList.Add((excelKey, errorExcelLog, excelName));
                            }
                            cellFix.Value = double.TryParse(cellFixValue, out double number) ? number : cellFixValue;
                        }
                    }
                });
            }
        });
        return errorList;
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

    public static (List<string>,int) ExcelDataWriteIdGroup(dynamic excelName, dynamic addValue, ExcelWorksheet sheet, dynamic fixKey, dynamic modelId)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var excelFileFixKey = 2;
        var writeIdList = new List<string>();
        int lastRow = 0;
        for (var excelMulti = 0; excelMulti < modelId[excelName].Count; excelMulti++)
        {
            var startValue = modelId[excelName][excelMulti].Item1[0, 0].ToString();
            var endValue = modelId[excelName][excelMulti].Item1[1, 0].ToString();
            var startRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, startValue);
            var endRowSource = ExcelDataAutoInsert.FindSourceRow(sheet, 2, endValue);
            string excelKeyMethod = fixKey[excelName][excelMulti].Item1[1, 0]?.ToString();
            //获取要写入的ID
            var count = endRowSource - startRowSource + 1;
            for (var i = 0; i < count; i++)
            {
                var cellSource = sheet.Cells[startRowSource + i, excelFileFixKey];
                if (cellSource.Value == null) continue;
                if (cellSource.Value.ToString() == "" || cellSource.Value.ToString() == "0") continue;
                //字段值改写方法
                var temp1 = ExcelDataAutoInsert.CellFixValueKeyList(excelKeyMethod);
                //修改字符串
                var cellFixValue = ExcelDataAutoInsert.StringRegPlace(cellSource.Value.ToString(), temp1, addValue);
                writeIdList.Add(cellFixValue);
            }
            //获取最后一行
            if (lastRow < endRowSource)
            {
                lastRow = endRowSource;
            }
        }
        return (writeIdList,lastRow);
    }
}