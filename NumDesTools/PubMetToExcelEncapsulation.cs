using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using Color = System.Drawing.Color;
using System.Threading.Tasks;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Data.OleDb;
using System.Diagnostics;
using ExcelReference = ExcelDna.Integration.ExcelReference;
using System.Text.RegularExpressions;
using Range = Microsoft.Office.Interop.Excel.Range;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using NumDesTools;

// ReSharper disable All
#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类-封装为类和属性
/// </summary>
public class Temp
{
    //使用类-属性-方法的封装案例

    //声明类的属性，封装方法计算的返回值为类的属性
    public List<object> SheetHeader { get; private set; }
    public List<List<object>> SheetData { get; private set; }
    //初始化属性
    public Temp()
    {
        SheetHeader = new List<object>();
        SheetData = new List<List<object>>();
    }
    //具体方法获取属性的值
    public void LoadDataFromRange(object[,] rangeValue)
    {
        var rows = rangeValue.GetLength(0);
        var columns = rangeValue.GetLength(1);

        for (var row = 1; row <= rows; row++)
        {
            var rowList = new List<object>();
            for (var column = 1; column <= columns; column++)
            {
                var value = rangeValue[row, column];
                if (row == 1)
                    SheetHeader.Add(value);
                else
                    rowList.Add(value);
            }

            if (row > 1) SheetData.Add(rowList);
        }
    }
}
public class TempTest
{
    // 使用方式：
    public static void TempList(dynamic workSheet)
    {
        Range dataRange = workSheet.UsedRange;
        object[,] rangeValue = dataRange.Value;

        //实例对象
        var excelWorksheet = new Temp();
        //调用方法计算属性
        excelWorksheet.LoadDataFromRange(rangeValue);
        //直接引用属性
        var sheetHeader = excelWorksheet.SheetHeader;
        var sheetData  = excelWorksheet.SheetData;
    }
}
