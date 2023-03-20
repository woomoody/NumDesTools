using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using static System.IO.Path;
using IHyperlink = NPOI.SS.UserModel.IHyperlink;

namespace NumDesTools;

public class AutoInsertData
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic indexWk = App.ActiveWorkbook;
    private static readonly dynamic indexWs = indexWk.Worksheets["索引关键词"];
    private static readonly dynamic Ws = indexWk.Worksheets["活动模板A类"];
    private static readonly dynamic WsTitle = indexWk.Worksheets["表头集合"];
    private static readonly dynamic autoKey = indexWs.Range["B2:B1000"];
    private static readonly object Missing = Type.Missing;
    private static readonly dynamic CacRowStart = 16; //角色参数配置行数起点

    //获取全部角色的关键数据（要导出的），生成List
    public static void GetExcelTitle()
    {
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;

        var mode = Ws.Range["B4"].Value2;
        var startExcel = Ws.Range["C4"].Value2;
        var startExcelIndex = Ws.Range["E4"].Value2;
        var endExcelIndex = Ws.Range["F4"].Value2;
        var dataMode = Ws.Range["D4"].Value2;


        var autokeyList = RangeToList(autoKey);
        string workPath = App.ActiveWorkbook.Path + @"\"+ startExcel;
        Workbook book = App.Workbooks.Open(workPath, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
        var ws2 = book.Worksheets[1];
        var endRow = ws2.UsedRange.Rows.Count;
        var rowRange = ws2.Range["B1:B"+ endRow];
        //需要自动修改的ID字段
        var keyColList = new List<string>();
        var endColumn = ws2.UsedRange.Columns.Count;
        var colRange = ws2.Range[ws2.Cells[1, 1], ws2.Cells[1, endColumn]];
        for (int i = 0; i < autokeyList.Count; i++)
        {
            var sourceCol = FindValueInColumns(colRange, autokeyList[i], 1);
            keyColList.AddRange(sourceCol);
        }
        //找到要Copy数据行
        var sourceRow=FindValueInRows(rowRange, dataMode);
        for (int i = 0; i <= endExcelIndex - startExcelIndex; i++)
        {
            var sheetId = ws2.Cells[i + 1 + sourceRow, 2].Value;
            if (sheetId != Convert.ToDouble(startExcelIndex + i) || sheetId==null )
            {
                //拷贝数据
                var sourceRange = ws2.Rows[sourceRow];
                var targetRange = ws2.Rows[sourceRow + 1+i];
                targetRange.Insert(XlInsertShiftDirection.xlShiftDown);
                var newTargetRange = ws2.Rows[sourceRow + 1+i];
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

    private static string RegNumReplaceNew(string text,int digit)
    {
        string pattern = "\\d+";
        // 使用正则表达式匹配数字
        MatchCollection matches = Regex.Matches(text, pattern);
        foreach (System.Text.RegularExpressions.Match match in matches)
        {
            string numStr = match.Value;
            int num = int.Parse(numStr);
            int newNum = num + (int)Math.Pow(10, digit - 1); // 对指定位数加1
            text = text.Replace(numStr, newNum.ToString());
        }
        return text;
    }

    public static List<string> RangeToList(dynamic range)
    {
        List<string> rangeList = new List<string>();
        foreach (Range r in range)
        {
            var tarV = r.Value2;
            if (tarV != null)
            {
                rangeList.Add(tarV.ToString());
            }
        }
        return rangeList;
    }
    public static int FindValueInRows(dynamic rowRange, dynamic dataMode)
    {
        Range row = null;
        //查找模板ID所在行
        foreach (Range r in rowRange)
        {
            var tarV = r.Value2;
            if (tarV != null)
            {
                if (tarV == dataMode.ToString())
                {
                    row = r;
                    break;
                }
            }
        }
        return row.Row;
    }
    public static List<string> FindValueInColumns(dynamic rowRange, dynamic dataMode,int mode)
    {
        List<string> colList = new List<string>();
        Range col = null;
        foreach (Range r in rowRange)
        {
            var tarV = r.Value2;
            if (tarV != null)
            {
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

    public static void GetExcelTitleNPOI2()
    {
        var asd =ExcelLinkExcel(@"D:\M1Work\public\Excels\Tables\");
        var fpe = @"D:\M1Work\public\Excels\Tables\#自动填表.xlsm";
        ExcelLinkGroup(asd, fpe);
    }

    public static void GetExcelTitleNPOI()
    {
        // 写入的工作簿
        var fpe = @"D:\M1Work\public\Excels\Tables\#自动填表.xlsm";
        var file = new FileStream(fpe, FileMode.Open, FileAccess.Read);
        // 创建工作簿对象
        var workbook = new XSSFWorkbook(file);
        // 获取第一个工作表
        // 如果工作簿中已经存在同名的工作表，则先删除该工作表
        if (workbook.GetSheet("表头集合") != null)
        {
            workbook.RemoveSheetAt(workbook.GetSheetIndex("表头集合"));
        }
        // 创建新的工作表
        ISheet sheet = workbook.CreateSheet("表头集合");
        //过滤空单元格
        var asd = sheet.LastRowNum;
        for (var i = 0; i <= asd; i++)
        {
            var row = (XSSFRow)sheet.GetRow(i);
            if (row == null) continue;
            var cell = (XSSFCell)row.GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            // 如果单元格为空，跳过该单元格
            if (cell.CellType == CellType.Blank) continue;
            var asd123 = cell.ToString();
        }


        // 读取的工作簿
        string sourceFilePath = @"D:\M1Work\public\Excels\Tables\";
        //var sfile = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read);
        //var sworkbook = new XSSFWorkbook(sfile);
        //var ssheet = sworkbook.GetSheetAt(0);


        string[] files = Directory.GetFiles(sourceFilePath,"*.xlsx");
        //string[] dirs = Directory.GetDirectories(path);
        var fcount = 0;
        // 遍历当前目录下所有文件
        foreach (string f in files)
        {
            // 判断文件名是否包含#
            if (!f.Contains("#"))
            {
                var sworkbook = new XSSFWorkbook(f);
                var ssheet = sworkbook.GetSheetAt(0);
                var fieName = Path.GetFileName(f);
                //sheet.GetRow(2*fcount).GetCell(0).SetCellValue(GetFileName(f));
                //第几行
                var srow = ssheet.GetRow(0) ?? ssheet.CreateRow(0);
                var row = sheet.GetRow(2*fcount) ?? sheet.CreateRow( 2 * fcount);
                var rowCol = sheet.GetRow(2 * fcount+1) ?? sheet.CreateRow(2 * fcount+1);
                var nameRow = sheet.GetRow(2 * fcount) ?? sheet.CreateRow(2 * fcount);
                var cellName = nameRow.GetCell(0) ?? nameRow.CreateCell(0);
                cellName.SetCellValue(fieName);
                var cellCount = 0;
                for (var j = 2; j < srow.LastCellNum; j++)
                {
                    //第几列
                    var scell = srow.GetCell(j) ?? srow.CreateCell(j);
                    var cell = row.GetCell(j- cellCount) ?? row.CreateCell(j- cellCount);
                    var cellCol = rowCol.GetCell(j- cellCount) ??rowCol.CreateCell(j- cellCount);
                    string sLink = null;
                    if (scell.Hyperlink != null)
                    {
                        sLink = scell.Hyperlink.Address;
                        switch (scell.CellType)
                        {
                        case CellType.Numeric:
                            cell.SetCellValue(scell.NumericCellValue);
                            cellCol.SetCellValue(j+1);
                            break;
                        case CellType.String:
                            cell.SetCellValue(scell.StringCellValue);
                            cellCol.SetCellValue(j + 1);
                            break;
                        case CellType.Boolean:
                            cell.SetCellValue(scell.BooleanCellValue);
                            cellCol.SetCellValue(j + 1);
                            break;
                        case CellType.Formula:
                            cell.CellFormula = scell.CellFormula;
                            cellCol.SetCellValue(j + 1);
                            break;
                        case CellType.Blank:
                            cellCount++;
                            break;
                        default:
                            break;
                        }
                    }
                    else
                    {
                        cellCount++;
                    }
                    if (sLink != null)
                    {
                        IHyperlink hyperlink = sheet.Workbook.GetCreationHelper().CreateHyperlink(HyperlinkType.Url);
                        hyperlink.Address = sLink;
                        cell.Hyperlink = hyperlink;
                    }
                }
                fcount++;
            }
        }
        var fileStream = new FileStream(fpe, FileMode.Create, FileAccess.Write);
        workbook.Write(fileStream);
        fileStream.Close();
        file.Close();
        workbook.Close();



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

    }

    private static List<List<(int, string, string,string, string)>> ExcelLinkExcel(string folderPath)
    {
        var excelTitleList = new List<List<(int, string, string, string, string)>>();
        // 读取的工作簿
        var files = Directory.GetFiles(folderPath, "*.xlsx");
        var fcount = 0;
        var fileNamesToCheck = new List<string>();
        foreach (var file in files) fileNamesToCheck.Add(GetFileName(file));
        var asdb = 0;
        // 遍历当前目录下所有文件
        foreach (var f in files)
        {
            var fileName = GetFileName(f);
            // 判断文件名是否包含#
            if (!f.Contains("#"))
            {
                var titleList = new List<(int, string, string, string, string)>();
                var fs = new FileStream(f, FileMode.Open, FileAccess.Read);
                var sworkbook = new XSSFWorkbook(fs);
                var ssheet = sworkbook.GetSheetAt(0);
                //第几行
                var srow = ssheet.GetRow(3) ?? ssheet.CreateRow(3);
                var row = ssheet.GetRow(0) ?? ssheet.CreateRow(0);
                for (var j = 0; j < srow.LastCellNum; j++)
                {
                    //第几列
                    var scell = srow.GetCell(j) ?? srow.CreateCell(j);
                    if (scell.CellComment != null)
                    {
                        var cell = row.GetCell(j);
                        if (cell == null)
                        {
                            cell = row.CreateCell(j);
                            cell.CellStyle = row.GetCell(1).CellStyle;
                        }

                        var newComment = scell.CellComment;
                        if (cell.CellType == CellType.Blank || cell.CellType == CellType.Unknown)
                        {
                            var linkFile = newComment.String.String;
                            if (linkFile == null) continue;
                            //var fileki ="在【ScoreTrigger.xlsx】表中配置对应类型数据";
                            // 匹配英文字母
                            var pattern = "[a-zA-Z]+";
                            // 创建正则表达式对象
                            var regex = new Regex(pattern);
                            // 查找匹配项
                            var matches = regex.Matches(linkFile);
                            foreach (Match match in matches)
                            {

                                // 判断文件名是否包含#
                                if (!f.Contains("#"))
                                {
                                    var isLinkExcel = fileNamesToCheck.Any(s =>
                                        string.Equals(s, match.Value + ".xlsx", StringComparison.OrdinalIgnoreCase));
                                    if (isLinkExcel)
                                    {
                                        var ad = j + 1;
                                        //cell.SetCellValue(match.Value);
                                        //var link = new XSSFHyperlink(HyperlinkType.File);
                                        //link.Address = sourceFilePath + match.Value + ".xlsx";
                                        //cell.Hyperlink = link;
                                        var cellValue = match.Value;
                                        var excelLink = folderPath + cellValue + ".xlsx";
                                        var cellLink =fileName+"#"+ssheet.SheetName+"!"+ssheet.GetRow(3).GetCell(j).Address;
                                        titleList.Add((j + 1, cellValue, excelLink,cellLink, fileName));
                                        //只取第一个索引
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                excelTitleList.Add(titleList);
                sworkbook.Close();
                fs.Close();
            }
        }
        return excelTitleList;
    }

    private static void ExcelLinkGroup(List<List<(int, string, string, string,string)>> excelTitleList,string filePath)
    {
        var fileCount = 0;
        var file2 = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        // 创建工作簿对象
        var workbook = new XSSFWorkbook(file2);
        // 获取第一个工作表
        // 如果工作簿中已经存在同名的工作表，则先删除该工作表
        if (workbook.GetSheet("表头集合") != null)
        {
            workbook.RemoveSheetAt(workbook.GetSheetIndex("表头集合"));
        }

        // 创建新的工作表
        ISheet sheet = workbook.CreateSheet("表头集合");
        for (int i = 0; i < excelTitleList.Count; i++)
        {
                if (excelTitleList[i].Count == 0)
                {
                    fileCount++;
                }
                else
                {
                    //第几行
                    var row = sheet.GetRow(+(i-fileCount) * 2) ?? sheet.CreateRow(+(i - fileCount) * 2);
                    var rowCOL = sheet.GetRow(1 + (i - fileCount) * 2) ?? sheet.CreateRow(1 + (i - fileCount) * 2);
                    for (int k = 0; k < excelTitleList[i].Count; k++)
                    {
                        var cell = row.GetCell(k + 1) ?? row.CreateCell(k + 1);
                        var cellCOL = rowCOL.GetCell(k + 1) ?? rowCOL.CreateCell(k + 1);
                        var cellExcelName = row.GetCell(0) ?? row.CreateCell(0);
                        cellExcelName.SetCellValue(excelTitleList[i][k].Item5);
                        cell.SetCellValue(excelTitleList[i][k].Item2);
                        var link = new XSSFHyperlink(HyperlinkType.File);
                        link.Address = excelTitleList[i][k].Item3;
                        cell.Hyperlink = link;
                        cellCOL.SetCellValue(excelTitleList[i][k].Item1);
                        var link2 = new XSSFHyperlink(HyperlinkType.File);
                        link2.Address = excelTitleList[i][k].Item4;
                        cellCOL.Hyperlink = link2;
                    }
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
        var path =indexWk.Path;
        var sheet = indexWk.Worksheets["活动模板A类"];
        var firstExcel = sheet.Range["C4"].Value;
        //初始表格的数据
        string workbookPath = path+@"\"+ firstExcel;
        var workBook = new XSSFWorkbook(workbookPath);
        var worSheet = workBook.GetSheetAt(0);
        var sheetIndexGroup = GetAnySheetRangeData(false,worSheet,1,1,1,0);
        var indexGroupSheet = indexWk.Worksheets("活动模板A类-索引");
        var jump = 0;
        for (int i = 0; i < sheetIndexGroup[0].Count; i++)
        {
            if (sheetIndexGroup[0][i].Item2 == "#" )
            {
                jump++;
            }
            else
            {
                if (sheetIndexGroup[0][i].Item2.Contains("*") || sheetIndexGroup[0][i].Item2.Contains(".xlsx"))
                {
                    indexGroupSheet.cells[i + 2 - jump, 1].value = sheetIndexGroup[0][i].Item2;
                }
                else
                {
                    jump++;
                }
            }
        }
    }
    //不知道多少行多少列就填0，会自动获取表格最大行列,true包含空值
    public static List<List<(int, string)>> GetAnySheetRangeData(bool isNull,dynamic workSheet, int rowStart, int colStart, int rowEnd ,int colEnd )
    {
        //获取数据边界
        var colMax = 0;
        var rowMax = workSheet.LastRowNum + 1;
        if (rowEnd == 0)
        {
            rowEnd = rowMax;
        }
        if (colEnd == 0)
        {
            for (int i = rowStart - 1; i < rowEnd; i++)
            {
                colMax = Math.Max(colMax, workSheet.GetRow(i).LastCellNum);
            }
            colEnd = colMax;
        }
        var rangeData = new List<List<(int,string)>>();
        for (int i = rowStart-1; i < rowEnd; i++)
        {
            var row = workSheet.GetRow(i);
            var cellData = new List<(int, string)>();
            for (int j = colStart - 1; j < colEnd; j++)
            {
                var cell= row.GetCell(j);
                if (isNull == true)
                {
                    if (cell == null)
                    {
                        cell = "";
                    }
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
    public static void ExcelIndexCircle()
    {
        App.DisplayAlerts = false;
        App.ScreenUpdating = false;
        var mode = Ws.Range["B4"].Value2;
        var startExcel = Ws.Range["C4"].Value2;
        var startExcelIndex = Ws.Range["E4"].Value2;
        var endExcelIndex = Ws.Range["F4"].Value2;
        var dataMode = Ws.Range["D4"].Value2;

        
        var titleExcelEndRow = WsTitle.UsedRange.Rows.Count;
        var titleExcelEndCol = WsTitle.UsedRange.Columns.Count;
        //展开母表所有关联表格
        var excelList= new List<string>();
        var titleExcelRangeRow = WsTitle.Range[WsTitle.Cells[1,1],WsTitle.Cells[titleExcelEndRow, 1]];
        var titleExcelRow =FindValueInRows(titleExcelRangeRow, startExcel);
        for (int j = 1; j < titleExcelEndCol + 1; j++)
        {
            var excelName = WsTitle.Cells[1, j + 1].Value;
            if (excelName != null)
            {
                excelList .Add( WsTitle.Cells[1, j + 1].Value + ".xlsx");
            }
        }
        var excelListAll = new List<string>();
        //展开字表所有关联表格
   
            for (int i = excelListAll.Count; i < excelList.Count ; i++)
            {
                for (int j = 0; j < titleExcelEndCol + 1; j++)
                {
                    var titleExcelRangeRowTemp = WsTitle.Range[WsTitle.Cells[1, 1], WsTitle.Cells[titleExcelEndRow, 1]];
                    var titleExcelRowTemp = FindValueInRows(titleExcelRangeRowTemp, excelList[i]);
                    var excelName = WsTitle.Cells[titleExcelRowTemp, j + 1].Value;
                    if (excelName != null)
                    {
                        excelListAll.Add(WsTitle.Cells[titleExcelRowTemp, j + 1].Value + ".xlsx");
                        excelList.Add(WsTitle.Cells[titleExcelRowTemp, j + 1].Value + ".xlsx");
                    }
                }
            }
            App.DisplayAlerts = true;
        App.ScreenUpdating = true;

    }
}
