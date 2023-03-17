using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace NumDesTools;

public class AutoInsertData
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic indexWk = App.ActiveWorkbook;
    private static readonly dynamic indexWs = indexWk.Worksheets["索引关键词"];
    private static readonly dynamic Ws = indexWk.Worksheets["设计表"];
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
        foreach (Match match in matches)
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
}
