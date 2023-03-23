using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming.Values;
using NPOI.XSSF.UserModel;
using static System.IO.Path;

namespace NumDesTools;

public class AutoInsertData
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic IndexWk = App.ActiveWorkbook;
    private static readonly dynamic IndexWs = IndexWk.Worksheets["索引关键词"];
    private static readonly dynamic Ws = IndexWk.Worksheets["活动模板A类"];
    private static readonly dynamic AutoKey = IndexWs.Range["B2:B1000"];
    private static readonly object Missing = Type.Missing;

    //获取全部角色的关键数据（要导出的），生成List
    public static void GetExcelTitle()
    {
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;

        var startExcel = Ws.Range["C4"].Value2;
        var startExcelIndex = Ws.Range["E4"].Value2;
        var endExcelIndex = Ws.Range["F4"].Value2;
        var dataMode = Ws.Range["D4"].Value2;


        var autokeyList = RangeToList(AutoKey);
        string workPath = App.ActiveWorkbook.Path + @"\" + startExcel;
        Workbook book = App.Workbooks.Open(workPath, Missing, Missing, Missing, Missing, Missing, Missing, Missing,
            Missing, Missing, Missing, Missing, Missing, Missing, Missing);
        var ws2 = book.Worksheets[1];
        var endRow = ws2.UsedRange.Rows.Count;
        var rowRange = ws2.Range["B1:B" + endRow];
        //需要自动修改的ID字段
        var keyColList = new List<string>();
        var endColumn = ws2.UsedRange.Columns.Count;
        var colRange = ws2.Range[ws2.Cells[1, 1], ws2.Cells[1, endColumn]];
        for (var i = 0; i < autokeyList.Count; i++)
        {
            var sourceCol = FindValueInColumns(colRange, autokeyList[i], 1);
            keyColList.AddRange(sourceCol);
        }

        //找到要Copy数据行
        var sourceRow = FindValueInRows(rowRange, dataMode);
        for (var i = 0; i <= endExcelIndex - startExcelIndex; i++)
        {
            var sheetId = ws2.Cells[i + 1 + sourceRow, 2].Value;
            if (sheetId != Convert.ToDouble(startExcelIndex + i) || sheetId == null)
            {
                //拷贝数据
                var sourceRange = ws2.Rows[sourceRow];
                var targetRange = ws2.Rows[sourceRow + 1 + i];
                targetRange.Insert(XlInsertShiftDirection.xlShiftDown);
                var newTargetRange = ws2.Rows[sourceRow + 1 + i];
                sourceRange.Copy(newTargetRange);
                //修改数据-创建ID
                ws2.Cells[i + 1 + sourceRow, 2].Value = startExcelIndex + i;
            }

            //修改数据-其他默认自动填写数据---同源引用应该也是一样的规则，所以在新表添加新数据时，需要去重写入数据
            var abc = "[1001,1101,1201,1301][1800]";
            //Regex regex = new Regex(@"\d+");
            //int max = regex.Matches(abc)
            //    .Cast<Match>()
            //    .Select(m => int.Parse(m.Value))
            //    .Max();
            //string output = Regex.Replace(abc, @"\d+", m => (int.Parse(m.Value) + 1).ToString());
            //Debug.Print(output);

            //准备数据
            abc = RegNumReplaceNew(abc, 2);
            Debug.Print(abc);
            //写入ws2


            //汇总引用表格List，汇总表格需要添加IDList，循环直到没有索引
        }

        //var values =roleDataRng.Value;
        //var comment = roleDataRng.Comment?.Text;


        //var range2 = Ws.Range[Ws.Cells[1, 1], Ws.Cells[4, endColumn]];
        //range2.Value = values;
        //range2.Comment?.Delete();
        //if (comment != null)
        //{
        //    range2.AddComment(comment);
        //}

        //// 保存并关闭工作簿
        //book.Save();
        //book.Close();

        //// 释放资源
        //Marshal.ReleaseComObject(range2);
        //Marshal.ReleaseComObject(ws2);
        //Marshal.ReleaseComObject(book);
        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
    }

    public static void GetExcelTitleNpoi2()
    {
        var asd = ExcelLinkExcel(@"D:\M1Work\public\Excels\Tables\");
        var fpe = @"D:\M1Work\public\Excels\Tables\#自动填表.xlsm";
        ExcelLinkGroup(asd, fpe);
    }

    private static List<List<(int, string, string, string, string, string)>> ExcelLinkExcel(string folderPath)
    {
        var excelTitleList = new List<List<(int, string, string, string, string, string)>>();
        // 读取的工作簿
        var files = Directory.GetFiles(folderPath, "*.xlsx");
        var fileNamesToCheck = new List<string>();
        foreach (var file in files) fileNamesToCheck.Add(GetFileName(file));
        // 遍历当前目录下所有文件
        foreach (var f in files)
        {
            var fileName = GetFileName(f);
            // 判断文件名是否包含#
            if (!f.Contains("#"))
            {
                var titleList = new List<(int, string, string, string, string,string)>();
                var fs = new FileStream(f, FileMode.Open, FileAccess.Read);
                var workbook = new XSSFWorkbook(fs);
                var sheets = workbook.GetSheetAt(0);
                //第几行
                var rows = sheets.GetRow(3) ?? sheets.CreateRow(3);
                var rowsKey = sheets.GetRow(1) ?? sheets.CreateRow(1);
                var row = sheets.GetRow(0) ?? sheets.CreateRow(0);
                for (var j = 0; j < rows.LastCellNum; j++)
                {
                    //第几列
                    var cells = rows.GetCell(j) ?? rows.CreateCell(j);
                    if (cells.CellComment != null)
                    {
                        var cell = row.GetCell(j);
                        if (cell == null)
                        {
                            cell = row.CreateCell(j);
                            cell.CellStyle = row.GetCell(1).CellStyle;
                        }
                        var newComment = cells.CellComment;
                        if (cell.CellType == CellType.Blank || cell.CellType == CellType.Unknown)
                        {
                            var linkFile = newComment.String.String;
                            if (linkFile == null) continue;
                            // 匹配英文字母
                            var pattern = "[a-zA-Z]+";
                            // 创建正则表达式对象
                            var regex = new Regex(pattern);
                            // 查找匹配项
                            var matches = regex.Matches(linkFile);
                            foreach (System.Text.RegularExpressions.Match match in matches)
                            {
                                // 判断文件名是否包含#
                                if (!f.Contains("#"))
                                {
                                    var isLinkExcel = fileNamesToCheck.Any(s =>
                                        string.Equals(s, match.Value + ".xlsx", StringComparison.OrdinalIgnoreCase));
                                    if (isLinkExcel)
                                    {
                                        var cellValue = match.Value;
                                        var excelLink = folderPath + cellValue + ".xlsx";
                                        var cellLink = fileName + "#" + sheets.SheetName + "!" +
                                                       sheets.GetRow(3).GetCell(j).Address;
                                        var cellsKey = rowsKey.GetCell(j).StringCellValue;
                                        titleList.Add((j + 1, cellValue, excelLink, cellLink, fileName, cellsKey));
                                        //只取第一个索引
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                excelTitleList.Add(titleList);
                workbook.Close();
                fs.Close();
            }
        }

        return excelTitleList;
    }

    private static void ExcelLinkGroup(List<List<(int, string, string, string, string,string)>> excelTitleList,
        string filePath)
    {
        var fileCount = 0;
        var file2 = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        // 创建工作簿对象
        var workbook = new XSSFWorkbook(file2);
        // 获取第一个工作表
        // 如果工作簿中已经存在同名的工作表，则先删除该工作表
        if (workbook.GetSheet("表头集合") != null) workbook.RemoveSheetAt(workbook.GetSheetIndex("表头集合"));

        // 创建新的工作表
        var sheet = workbook.CreateSheet("表头集合");
        for (var i = 0; i < excelTitleList.Count; i++)
            if (excelTitleList[i].Count == 0)
            {
                fileCount++;
            }
            else
            {
                //第几行
                var row = sheet.GetRow(+(i - fileCount) * 3) ?? sheet.CreateRow(+(i - fileCount) * 3);
                var rowCol = sheet.GetRow(1 + (i - fileCount) * 3) ?? sheet.CreateRow(1 + (i - fileCount) * 3);
                var rowKey = sheet.GetRow(2 + (i - fileCount) * 3) ?? sheet.CreateRow(2 + (i - fileCount) * 3);
                for (var k = 0; k < excelTitleList[i].Count; k++)
                {
                    var cell = row.GetCell(k + 1) ?? row.CreateCell(k + 1);
                    var cellCol = rowCol.GetCell(k + 1) ?? rowCol.CreateCell(k + 1);
                    var cellKey =rowKey.GetCell(k + 1) ?? rowKey.CreateCell(k + 1);
                    var cellExcelName = row.GetCell(0) ?? row.CreateCell(0);
                    cellExcelName.SetCellValue(excelTitleList[i][k].Item5);
                    cell.SetCellValue(excelTitleList[i][k].Item2);
                    var link = new XSSFHyperlink(HyperlinkType.File)
                    {
                        Address = excelTitleList[i][k].Item3
                    };
                    cell.Hyperlink = link;
                    cellCol.SetCellValue(excelTitleList[i][k].Item1);
                    var link2 = new XSSFHyperlink(HyperlinkType.File)
                    {
                        Address = excelTitleList[i][k].Item4
                    };
                    cellCol.Hyperlink = link2;
                    cellKey.SetCellValue(excelTitleList[i][k].Item6);
                }
            }

        var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        workbook.Write(fileStream);
        workbook.Close();
        file2.Close();
        fileStream.Close();
    }

    public static void ExcelIndexGroup()
    {
        var path = IndexWk.Path;
        var sheet = IndexWk.Worksheets["活动模板A类"];
        var firstExcel = sheet.Range["C4"].Value;
        //初始表格的数据
        string workbookPath = path + @"\" + firstExcel;
        var workBook = new XSSFWorkbook(workbookPath);
        var worSheet = workBook.GetSheetAt(0);
        var sheetIndexGroup = GetAnySheetRangeData(false, worSheet, 1, 1, 1, 0);
        var indexGroupSheet = IndexWk.Worksheets("活动模板A类-索引");
        var jump = 0;
        for (var i = 0; i < sheetIndexGroup[0].Count; i++)
            if (sheetIndexGroup[0][i].Item2 == "#")
            {
                jump++;
            }
            else
            {
                if (sheetIndexGroup[0][i].Item2.Contains("*") || sheetIndexGroup[0][i].Item2.Contains(".xlsx"))
                    indexGroupSheet.cells[i + 2 - jump, 1].value = sheetIndexGroup[0][i].Item2;
                else
                    jump++;
            }
    }

    public static void ExcelIndexCircle()
    {
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        var wsTitle = IndexWk.Worksheets["表头集合"];
        var startExcel = Ws.Range["C4"].Value2;


        var titleExcelEndRow = wsTitle.UsedRange.Rows.Count;
        var titleExcelEndCol = wsTitle.UsedRange.Columns.Count;
        //展开母表所有关联表格
        var excelList = new List<string>();
        var titleExcelRangeRow = wsTitle.Range[wsTitle.Cells[1, 1], wsTitle.Cells[titleExcelEndRow, 1]];
        FindValueInRows(titleExcelRangeRow, startExcel);
        for (var j = 1; j < titleExcelEndCol + 1; j++)
        {
            var excelName = wsTitle.Cells[1, j + 1].Value;
            if (excelName != null) excelList.Add(wsTitle.Cells[1, j + 1].Value + ".xlsx");
        }

        var excelListAll = new List<string>();
        //展开字表所有关联表格

        for (var i = excelListAll.Count; i < excelList.Count; i++)
        for (var j = 0; j < titleExcelEndCol + 1; j++)
        {
            var titleExcelRangeRowTemp = wsTitle.Range[wsTitle.Cells[1, 1], wsTitle.Cells[titleExcelEndRow, 1]];
            var titleExcelRowTemp = FindValueInRows(titleExcelRangeRowTemp, excelList[i]);
            var excelName = wsTitle.Cells[titleExcelRowTemp, j + 1].Value;
            if (excelName != null)
            {
                excelListAll.Add(wsTitle.Cells[titleExcelRowTemp, j + 1].Value + ".xlsx");
                excelList.Add(wsTitle.Cells[titleExcelRowTemp, j + 1].Value + ".xlsx");
            }
        }

        App.DisplayAlerts = true;
        App.ScreenUpdating = true;
    }

    public static void Factorial()
    {
        Worksheet sheet =IndexWk.ActiveSheet;
        //读取模板表数据
        Dictionary<string, List<string>> relationships = new Dictionary<string, List<string>>();
        for (int i = 1; i <= 3; i++)
        {
            var baseValue = sheet.Cells[i, 1].Value.ToString();
            relationships[baseValue] = new List<string>();
            for (int j = 2; j <= 3; j++)
            {
                var linkValue = sheet.Cells[i,j].Value;
                relationships[baseValue].Add(linkValue);
            }
        }
        //备份字典
        Dictionary<string, List<string>> relationshipsBF = new Dictionary<string, List<string>>();
        foreach (var key in relationships.Keys)
        {
            relationshipsBF[key] = new List<string>(relationships[key]);
        }
        //整理文件关联字典
        Dictionary<string, List<string>> relationships2 = new Dictionary<string, List<string>>(relationships);
        Dictionary<string, List<string>> relationships3 = new Dictionary<string, List<string>>();
        var keys = relationships.Keys;
        foreach (var item in keys)
        {
            relationships3[item] = new List<string>();
            //初始表格映射
            foreach (var item2 in relationships["索引1.xlsx"])
            {
                relationships3[item].Add(item);
            }
            for (int i = 0; i < relationships[item].Count; i++)
            {
                var values = relationships[item][i];
                if (values == null) continue;
                if (relationships.ContainsKey(values))
                {
                    int asdb = 0;
                    foreach (var sdsa in relationships[values])
                    {
                        relationships2[item].Add(relationships[values][asdb]);
                        asdb++;
                        relationships3[item].Add(values);
                    }
                }
            }
        }

        //根据key[0]的文件流（包含所有文件，并且有既定的顺序）进行内容写入
        //var rootPath = IndexWk.Path;
        //var fileDic = relationships2["索引1.xlsx"];
        //for (int i = 0; i < fileDic.Count; i++)
        //{
        //    if (fileDic[i] == null) continue;//空的只是更改值，不会再索引表了
        //    var filePath = rootPath + @"\" + fileDic[i];
        //    var file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        //    var workbook = new XSSFWorkbook(file);
        //    var sheet1 = workbook.GetSheetAt(0);
        //    var row = sheet1.GetRow(0) ?? sheet1.CreateRow(0);
        //    var cell = row.GetCell(0) ?? row.CreateCell(0);
        //    cell.SetCellValue("abc");
        //    var file2 = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        //    workbook.Write(file2);
        //    workbook.Close();
        //    file.Close();
        //    file2.Close();
        //}
        bool areEqual = relationshipsBF.OrderBy(x => x.Key)
            .SequenceEqual(relationships2.OrderBy(x => x.Key));
        Debug.Print(areEqual.ToString());

    }

    public static void ActiveWorkbookWRDataByNPOI()
    {
        App.ActiveWorkbook.Close();
        var Missing = Type.Missing;

        var file2 = new FileStream(@"D:\M1Work\public\Excels\Tables\#自动填表.xlsm", FileMode.Open, FileAccess.Read);
        // 创建工作簿对象
        var workbook = new XSSFWorkbook(file2);
        // 获取第一个工作表
        // 如果工作簿中已经存在同名的工作表，则先删除该工作表
        var sheet = workbook.GetSheetAt(0);
        var row = sheet.GetRow(0) ?? sheet.CreateRow(0);
        var cell = row.GetCell(0) ?? row.CreateCell(0);
        cell.SetCellValue("123");
        var row2 = sheet.GetRow(1) ?? sheet.CreateRow(1);
        var cell2 = row2.GetCell(0) ?? row2.CreateCell(0);
        cell2.SetCellValue(cell.StringCellValue);
        var file3 = new FileStream(@"D:\M1Work\public\Excels\Tables\#自动填表.xlsm", FileMode.Create, FileAccess.Write);
        workbook.Write(file3);
        file2.Close();
        file3.Close();
        Workbook book = App.Workbooks.Open(@"D:\M1Work\public\Excels\Tables\#自动填表.xlsm", Missing, Missing, Missing,
            Missing, Missing, Missing, Missing,
            Missing, Missing, Missing, Missing, Missing, Missing, Missing);
    }

    //不知道多少行多少列就填0，会自动获取表格最大行列,true包含空值
    private static List<List<(int, string)>> GetAnySheetRangeData(bool isNull, dynamic workSheet, int rowStart,
        int colStart, int rowEnd, int colEnd)
    {
        //获取数据边界
        var colMax = 0;
        var rowMax = workSheet.LastRowNum + 1;
        if (rowEnd == 0) rowEnd = rowMax;
        if (colEnd == 0)
        {
            for (var i = rowStart - 1; i < rowEnd; i++) colMax = Math.Max(colMax, workSheet.GetRow(i).LastCellNum);
            colEnd = colMax;
        }

        var rangeData = new List<List<(int, string)>>();
        for (var i = rowStart - 1; i < rowEnd; i++)
        {
            var row = workSheet.GetRow(i);
            var cellData = new List<(int, string)>();
            for (var j = colStart - 1; j < colEnd; j++)
            {
                var cell = row.GetCell(j);
                if (isNull)
                {
                    if (cell == null) cell = "";
                    cellData.Add((j, cell.ToString()));
                }
                else
                {
                    if (cell == null) continue;
                    cellData.Add((j, cell.ToString()));
                }
            }

            rangeData.Add(cellData);
        }

        return rangeData;
    }

    private static string RegNumReplaceNew(string text, int digit)
    {
        var pattern = "\\d+";
        // 使用正则表达式匹配数字
        var matches = Regex.Matches(text, pattern);
        foreach (System.Text.RegularExpressions.Match match in matches)
        {
            var numStr = match.Value;
            var num = int.Parse(numStr);
            var newNum = num + (int)Math.Pow(10, digit - 1); // 对指定位数加1
            text = text.Replace(numStr, newNum.ToString());
        }

        return text;
    }

    private static List<string> RangeToList(dynamic range)
    {
        var rangeList = new List<string>();
        foreach (Range r in range)
        {
            var tarV = r.Value2;
            if (tarV != null) rangeList.Add(tarV.ToString());
        }

        return rangeList;
    }

    private static int FindValueInRows(dynamic rowRange, dynamic dataMode)
    {
        Range row = null;
        //查找模板ID所在行
        foreach (Range r in rowRange)
        {
            var tarV = r.Value2;
            if (tarV != null)
                if (tarV == dataMode.ToString())
                {
                    row = r;
                    break;
                }
        }

        Debug.Assert(row != null, nameof(row) + " != null");
        return row.Row;
    }

    private static List<string> FindValueInColumns(dynamic rowRange, dynamic dataMode, int mode)
    {
        var colList = new List<string>();
        foreach (Range r in rowRange)
        {
            var tarV = r.Value2;
            if (tarV != null)
            {
                Range col;
                if (mode == 0)
                {
                    if (tarV == dataMode.ToString())
                    {
                        col = r;
                        colList.Add(col.Column.ToString());
                    }
                }
                else if (mode == 1)
                {
                    if (tarV.Contains(dataMode.ToString()))
                    {
                        col = r;
                        colList.Add(col.Column.ToString());
                    }
                }
            }
        }

        return colList;
    }
}