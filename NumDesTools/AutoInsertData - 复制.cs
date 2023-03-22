using System;
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

public class AutoInsertData2
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic IndexWk = App.ActiveWorkbook;

    //a单元格的关联表格是b和c、b又关联d，c又关联e，e又关联b，展开a所有的关联
    public static void GetRelatedCell()
    {
        var activeCell = App.ActiveCell;
        var list = new List<string>();
        GetRelatedCell(activeCell, list);
        if (list.Count > 0)
        {
            App.Intersect(list.Select(x => IndexWk.Range[x]).ToArray()).Select();
        }
    }
    //获取单元格的关联单元格
    private static void GetRelatedCell(dynamic cell, List<string> list)
    {
        var formula = cell.Formula;
        if (string.IsNullOrEmpty(formula))
        {
            return;
        }
        var match = Regex.Match(formula, @"(?<=\{)(.*?)(?=\})");
        if (!match.Success)
        {
            return;
        }
        var cellName = match.Value;
        var cellRange = IndexWk.Range[cellName];
        if (cellRange.Count > 1)
        {
            foreach (var item in cellRange)
            {
                GetRelatedCell(item, list);
            }
        }
        else
        {
            GetRelatedCell(cellRange, list);
        }
        if (list.Contains(cellName))
        {
            return;
        }
        list.Add(cellName);
    }
    public static void AutoInsertData()
    {
        var activeCell = App.ActiveCell;
        var formula = activeCell.Formula;
        if (string.IsNullOrEmpty(formula))
        {
            return;
        }
        var match = Regex.Match(formula, @"(?<=\{)(.*?)(?=\})");
        if (!match.Success)
        {
            return;
        }
        var cellName = match.Value;
        var cellRange = IndexWk.Range[cellName];
        var cellValue = cellRange.Value;
        if (cellRange.Count > 1)
        {
            var list = new List<string>();
            foreach (var item in cellRange)
            {
                list.Add(item.Value);
            }
            activeCell.Value = string.Join(",", list);
        }
        else
        {
            activeCell.Value = cellValue;
        }
    }

}