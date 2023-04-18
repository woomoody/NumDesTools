using Microsoft.Office.Interop.Excel;
using System;
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
    public static Dictionary<string, List<Tuple<object[,]>>> ExcelDataToDictionary(dynamic data, dynamic dicKeyCol, dynamic dicValueCol,int valueRowCount, int valueColCount=1)
    {
        var dic = new Dictionary<string, List<Tuple<object[,]>>>();

        for (var i = 0; i < data.Count; i++)
        {
            var value = data[i][dicKeyCol];

            if (value == null) continue;

            object[,] values = new object[valueRowCount, valueColCount];
            for (int k = 0; k < valueRowCount; k++)
            {
                for (int j = 0; j < valueColCount; j++)
                {
                    var valueTemp = data[i + k][dicValueCol + j];
                    values[k, j] = valueTemp;
                }
            }
            var tuple = new Tuple<object[,]>(values);
            if (dic.ContainsKey(value))
            {
                dic[value].Add(tuple);
            }
            else
            {
                var list = new List<Tuple<object[,]>>();
                list.Add(tuple);
                dic.Add(value, list);
            }
        }
        return dic;
    }
}