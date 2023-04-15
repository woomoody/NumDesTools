using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace NumDesTools;

public class PubMetToExcel
{
    public static (List<object> sheetHeaderCol, List<List<object>> sheetData) ExcelDataToList(dynamic workSheet)
    {
        Range dataRange = workSheet.UsedRange;
        // 读取数据到一个二维数组中
        object[,] rangeValue = dataRange.Value;
        // 获取行数和列数
        int rows = rangeValue.GetLength(0);
        int columns = rangeValue.GetLength(1);
        // 定义工作表数据数组和表头数组
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        var sheetHeaderRow = new List<object>();
        // 读取数据和表头
        //单线程
        for (int row = 1; row <= rows; row++)
        {
            var rowList = new List<object>();
            for (int column = 1; column <= columns; column++)
            {
                object value = rangeValue[row, column];
                if (row == 1)
                {
                    sheetHeaderCol.Add(value);
                }
                else
                {
                    rowList.Add(value);
                }
            }

            if (row > 1)
            {
                sheetData.Add(rowList);
            }
        }
        (List<object> sheetHeaderCol, List<List<object>> sheetData) excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }
}