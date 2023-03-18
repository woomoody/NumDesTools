using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace NumDesTools;

public class AutoInsertData
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic indexWk = App.ActiveWorkbook;
    private static readonly dynamic indexWs = indexWk.Worksheets["索引关键词"];
    private static readonly dynamic Ws = indexWk.Worksheets["活动模板A类"];
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
    public static void GetExcelTitleNPOI()
    {

 
        // 写入的工作簿
        var fpe = @"D:\M1Work\public\Excels\Tables\#自动填表.xlsm";
        var file = new FileStream(fpe, FileMode.Open, FileAccess.Read);
        // 创建工作簿对象
        var workbook = new XSSFWorkbook(file);
        // 获取第一个工作表
        var sheet = workbook.GetSheet("表头集合");
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


        string[] files = Directory.GetFiles(sourceFilePath);
        //string[] dirs = Directory.GetDirectories(path);
        var fcount = 0;
        // 遍历当前目录下所有文件
        foreach (string f in files)
        {
            // 判断文件名是否包含#
            if (!f.Contains("#"))
            {
                // 判断是否为.xlsx文件
                if (Path.GetExtension(f) == ".xlsx")
                {
                    var sworkbook = new XSSFWorkbook(f);
                    var ssheet = sworkbook.GetSheetAt(0);
                    for (var i = 0; i < 4; i++)
                    {
                        //第几行
                        var srow = ssheet.GetRow(i) ?? ssheet.CreateRow(i);
                        var row = sheet.GetRow(i+4*fcount) ?? sheet.CreateRow(i+ 4 * fcount);
                        for (var j = 0; j < srow.LastCellNum; j++)
                        {
                            //第几列
                            var scell = srow.GetCell(j) ?? srow.CreateCell(j);
                            var cell = row.GetCell(j) ?? row.CreateCell(j);
                            switch (scell.CellType)
                            {
                                case CellType.Numeric:
                                    cell.SetCellValue(scell.NumericCellValue);
                                    break;
                                case CellType.String:
                                    cell.SetCellValue(scell.StringCellValue);
                                    break;
                                case CellType.Boolean:
                                    cell.SetCellValue(scell.BooleanCellValue);
                                    break;
                                case CellType.Formula:
                                    cell.CellFormula = scell.CellFormula;
                                    break;
                                case CellType.Blank:
                                    // do nothing
                                    break;
                                default:
                                    // do nothing
                                    break;
                            }
                            // 检查源单元格是否有注释，如果有则将注释添加到新单元格的批注中
                            //if (scell.CellComment != null)
                            //{
                            //    string author = scell.CellComment.Author;
                            //    string comment = scell.CellComment.String.String;
                            //    NPOI.SS.UserModel.IComment newComment = cell.CellComment;
                            //    if (newComment == null)
                            //    {
                            //        newComment = scell.CellComment = ssheet.CreateDrawingPatriarch().CreateCellComment(new XSSFClientAnchor());
                            //    }
                            //    newComment.Author = author;
                            //    newComment.String = scell.CellComment.String;
                            //}
                            if (scell.CellComment != null)
                            {
                                //NPOI.SS.UserModel.IComment sourceComment = scell.CellComment;
                                //NPOI.SS.UserModel.IComment newComment = cell.CellComment;
                                //newComment.Author = sourceComment.Author;

                                //// 为注释添加富文本字符串
                                //IRichTextString rt = sourceComment.String;
                                //newComment.String = rt;

                                //// 关联注释与新单元格对象
                                //newComment.c = cell;
                                //if (cell.CellComment == null)
                                //{
                                //    cell.CellComment = scell.CellComment = ssheet.CreateDrawingPatriarch().CreateCellComment(new XSSFClientAnchor());
                                //}
                                //cell.CellComment.Visible = scell.CellComment.Visible;
                                //cell.CellComment.String = scell.CellComment.String;
                            }
                        }
                    }
                    fcount++;
                }
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
}
