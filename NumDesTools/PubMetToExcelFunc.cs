using ExcelDna.Integration;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System;
using System.Diagnostics;
using DocumentFormat.OpenXml.Math;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace NumDesTools;
/// <summary>
/// 公共的Excel功能类调用的具体业务逻辑
/// </summary>
public class PubMetToExcelFunc
{
    private static readonly dynamic Wk = CreatRibbon._app.ActiveWorkbook;
    private static readonly dynamic Path = Wk.Path;
    //Excel数据查询并合并表格数据
    public static void ExcelDataSearchAndMerge(string searchValue)
    {
        //获取所有的表格路径
        string[] ignoreFileNames = { "#","副本"};
        var rootPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(Path));
        var fileList = new List<string>() { rootPath+ @"\Excels\Tables\", rootPath + @"\Excels\Localizations\", rootPath + @"\Excels\UIs\" };
        var files = PubMetToExcel.PathExcelFileCollect(fileList, "*.xlsx", ignoreFileNames);
        //查找指定关键词，记录行号和表格索引号
        var findValueList = new List<(string, string, int, int,string,string)>();
        Parallel.ForEach(files, file =>
        {
            var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(file);
            var findValue = PubMetToExcel.FindDataInDataTable(file , dataTable, searchValue);
            if (findValue.Count > 0)
            {
                //findValueList.Add(findValue);
                findValueList = findValueList.Concat(findValue).ToList();
            }
        }
            );
        //人工查询所需要的数据，可以打开表格，可以删除和手动增加数据，专用表格进行操作
        dynamic tempWorkbook;
        try
        {
            tempWorkbook = CreatRibbon._app.Workbooks.Open(rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
        }
        catch
        {
            tempWorkbook = CreatRibbon._app.Workbooks.Add();
            tempWorkbook.SaveAs(rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
        }
        dynamic tempSheet = tempWorkbook.Sheets["Sheet1"];
        string[,] tempDataArray = new string[findValueList.Count, 5];
        for (int i = 0; i < findValueList.Count; i++)
        {
            tempDataArray[i, 0] = findValueList[i].Item1;
            tempDataArray[i, 1] = findValueList[i].Item2;
            tempDataArray[i, 2] = PubMetToExcel.ConvertToExcelColumn(findValueList[i].Item4)+findValueList[i].Item3;
            tempDataArray[i, 3] = findValueList[i].Item5;
            tempDataArray[i, 4] = findValueList[i].Item6;
            
        }
        var tempDataRange = tempSheet.Range[tempSheet.Cells[2, 2], tempSheet.Cells[2 + tempDataArray.GetLength(0) - 1, 2 + tempDataArray.GetLength(1) - 1]];
        tempDataRange.Value = tempDataArray;
        tempWorkbook.Save();
        //合并数据
    }
    //Excel右键识别文件路径并打开
    public static void RightOpenExcelByActiveCell(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sheet = CreatRibbon._app.ActiveSheet;
        var selectCell = CreatRibbon._app.ActiveCell;
        string selectCellValue = "";
        if (selectCell.Value != null)
        {
            selectCellValue = selectCell.Value.ToString();
        }
        //正则出是Excel路径的单元格
        var isMatch = Regex.IsMatch(selectCellValue, @"^[A-Za-z]:(\\[\w-]+)+(\.xlsx)$");
        if (isMatch)
        {
            var selectRow = selectCell.Row;
            var selectCol = selectCell.Column;
            var sheetName = sheet.Cells[selectRow, selectCol+1].Value;
            var cellAdress = sheet.Cells[selectRow, selectCol + 2].Value;
            PubMetToExcel.OpenExcelAndSelectCell(selectCellValue,sheetName,cellAdress);
        }


    }

    public static void AliceBigRicherDFS()
    {
        var ws = Wk.ActiveSheet;
        string targetA = ws.Range["A3"].Value.ToString();
        string targetB = ws.Range["B3"].Value.ToString();
        string targetC = ws.Range["C3"].Value.ToString();
        string targetCount =ws.Range["D3"].Value.ToString();
        var filterDataRange = ws.Range["E14:AE1454"];
        // 读取数据到一个二维数组中
        object[,] filterDataRangeValue = filterDataRange.Value;
        var filterDataRangeValueList = PubMetToExcel.RangeDataToList(filterDataRangeValue);

        // 使用正则表达式匹配数字
        var numbersA = Regex.Split(targetA, "#");
        var numbersB = Regex.Split(targetB, "#");
        var numbersC = Regex.Split(targetC, "#");
        var numbersCount = Regex.Split(targetCount, "#");
        // 使用LINQ进行筛选
        List<List<object>> filteredRows = filterDataRangeValueList
            .Where(row => row[Convert.ToInt32(numbersA[0])+8].ToString() == numbersA[1])
            .ToList();
        if (numbersB[0] != "")
        {
            filteredRows = filteredRows
                .Where(row => row[Convert.ToInt32(numbersB[0]) + 8].ToString() == numbersB[1])
                .ToList();
        }
        if (numbersC[0] != "0")
        {
            filteredRows = filteredRows
                .Where(row => row[Convert.ToInt32(numbersC[0]) + 8].ToString() == numbersC[1])
                .ToList();
        }
        if (numbersCount[0] != "0")
        {
            filteredRows = filteredRows
                    .Where(row => numbersCount.Any(condition => row[25].ToString() == condition))
                    .ToList();
        }
        int columnIndex = 26; // 第四列（索引从0开始）
        var errorLog="";
        // 写入每一行的指定列数据
        foreach (var row in filteredRows)
        {
            errorLog+= row[columnIndex]+"\n";
        }
        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(errorLog);

        // 释放 COM 对象
        Marshal.ReleaseComObject(ws);
        Marshal.ReleaseComObject(Wk);
        Marshal.ReleaseComObject(CreatRibbon._app);

        //int targetSum = (int)a;
        //int numberOfNumbers = (int)b;
        //int maxnum = (int)c;

        //for (int i = numberOfNumbers; i <= maxnum; i++)
        //{
        //    List<List<int>> combinations = FindCombinations(targetSum, i);
        //    foreach (var combination in combinations)
        //    {
        //        Debug.Print(string.Join(", ", combination));
        //    }
        //}



        //int[] numbers = { 1, 2, 3, 4 ,5, 6};


        //Debug.Print("所有可能的随机排序：");
        //EnumeratePermutations(numbers, 0, numbers.Length - 1);
    }

    static void EnumeratePermutations(int[] array, int startIndex, int endIndex)
    {
        if (startIndex == endIndex)
        {
            // 打印当前排列
            PrintArray(array);
        }
        else
        {
            for (int i = startIndex; i <= endIndex; i++)
            {
                // 交换元素，生成不同的排列
                Swap(ref array[startIndex], ref array[i]);

                // 递归调用，生成下一个位置的排列
                EnumeratePermutations(array, startIndex + 1, endIndex);

                // 恢复数组，以便后续交换
                Swap(ref array[startIndex], ref array[i]);
            }
        }
    }

    static void Swap(ref int a, ref int b)
    {
        int temp = a;
        a = b;
        b = temp;
    }

    static void PrintArray(int[] array)
    {
        string abc="";
        foreach (var number in array)
        {
            
           abc += number.ToString()+",";
        }
        Debug.Print(abc);
    }

    public static void TestCAPI()
    {
        Wk.Cells[1,1].Value = "Hello World";



    }

}
