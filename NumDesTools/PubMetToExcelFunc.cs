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
using NPOI.SS.Util;
using OfficeOpenXml;

namespace NumDesTools;
/// <summary>
/// 公共的Excel功能类调用的具体业务逻辑
/// </summary>
public class PubMetToExcelFunc
{
    private static readonly dynamic Wk = CreatRibbon._app.ActiveWorkbook;
    private static readonly dynamic path = Wk.Path;
    //Excel数据查询并合并表格数据
    public static void ExcelDataSearchAndMerge(string searchValue)
    {
        //获取所有的表格路径
        string[] ignoreFileNames = { "#","副本"};
        var rootPath = Path.GetDirectoryName(Path.GetDirectoryName(path));
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
    public static void OpenBaseLanExcel(CommandBarButton ctrl, ref bool cancelDefault)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var selectCell = CreatRibbon._app.ActiveCell;
        var basePath = CreatRibbon._app.ActiveWorkbook.Path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(basePath));
        newPath = newPath + @"\Excels\Localizations\Localizations.xlsx";
        var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(newPath);
        var findValue = PubMetToExcel.FindDataInDataTable(newPath, dataTable, selectCell.Value.ToString());
        var cellAddress = PubMetToExcel.ConvertToExcelColumn(findValue[0].Item4) + findValue[0].Item3;
        PubMetToExcel.OpenExcelAndSelectCell(newPath, "Sheet1", cellAddress);
    }
    public static void OpenMergeLanExcel(CommandBarButton ctrl, ref bool cancelDefault)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var selectCell = CreatRibbon._app.ActiveCell;
        var basePath = CreatRibbon._app.ActiveWorkbook.Path;
        string mergePath = "";
        //数据源路径txt
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        //第一行Alice，第二行Cove
        if (mergePathList.Count <= 1)
        {
            //打开文本文件
            Process.Start(filePath);
        }
        if (mergePathList[0] == "" || mergePathList[1] == "" || mergePathList[1] == mergePathList[0])
        {
            //打开文本文件
            Process.Start(filePath);
        }
        else
        {
            mergePath = basePath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        }
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(mergePath));
        newPath = newPath + @"\Excels\Localizations\Localizations.xlsx";
        var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(newPath);
        var findValue = PubMetToExcel.FindDataInDataTable(newPath, dataTable, selectCell.Value.ToString());
        string cellAddress = "";
        if (findValue.Count == 0)
        {
            cellAddress = "A1";
        }
        else
        {
            cellAddress = PubMetToExcel.ConvertToExcelColumn(findValue[0].Item4) + findValue[0].Item3;
        }
        PubMetToExcel.OpenExcelAndSelectCell(newPath, "Sheet1", cellAddress);
    }
    public static void AliceBigRicherDFS()
    {
        var ws = Wk.ActiveSheet;
        object[,] targetRank  = ws.Range["D11:D34"].Value;
        var targetRankList = PubMetToExcel.RangeDataToList(targetRank);
        object[,] seedRangeValue = ws.Range["D2:I2"].Value;
        var dataSeed = PubMetToExcel.RangeDataToList(seedRangeValue);
        object[,] targetKey = ws.Range["D4:E8"].Value;
        var targetKeyList = PubMetToExcel.RangeDataToList(targetKey);
        int maxRoll = Convert.ToInt32(ws.Range["D9"].Value);
        //string targetA = ws.Range["A3"].Value.ToString();
        //string targetB = ws.Range["B3"].Value.ToString();
        //string targetC = ws.Range["C3"].Value.ToString();
        //string targetCount =ws.Range["D3"].Value.ToString();
        //var filterDataRange = ws.Range["E14:AE1454"];
        //// 读取数据到一个二维数组中
        //object[,] filterDataRangeValue = filterDataRange.Value;
        //var filterDataRangeValueList = PubMetToExcel.RangeDataToList(filterDataRangeValue);

        //// 使用正则表达式匹配数字
        //var numbersA = Regex.Split(targetA, "#");
        //var numbersB = Regex.Split(targetB, "#");
        //var numbersC = Regex.Split(targetC, "#");
        //var numbersCount = Regex.Split(targetCount, "#");
        //// 使用LINQ进行筛选
        //List<List<object>> filteredRows = filterDataRangeValueList
        //    .Where(row => row[Convert.ToInt32(numbersA[0])+8].ToString() == numbersA[1])
        //    .ToList();
        //if (numbersB[0] != "")
        //{
        //    filteredRows = filteredRows
        //        .Where(row => row[Convert.ToInt32(numbersB[0]) + 8].ToString() == numbersB[1])
        //        .ToList();
        //}
        //if (numbersC[0] != "0")
        //{
        //    filteredRows = filteredRows
        //        .Where(row => row[Convert.ToInt32(numbersC[0]) + 8].ToString() == numbersC[1])
        //        .ToList();
        //}
        //if (numbersCount[0] != "0")
        //{
        //    filteredRows = filteredRows
        //            .Where(row => numbersCount.Any(condition => row[25].ToString() == condition))
        //            .ToList();
        //}
        //int columnIndex = 26; // 第四列（索引从0开始）
        //var errorLog="";
        //// 写入每一行的指定列数据
        //foreach (var row in filteredRows)
        //{
        //    errorLog+= row[columnIndex]+"\n";
        //}
        //ErrorLogCtp.DisposeCtp();
        //ErrorLogCtp.CreateCtpNormal(errorLog);



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
        List<int> data = new List<int>();
        for (int i = 0; i < dataSeed[0].Count; i++)
        {
            for (int j = 0; j < Convert.ToInt32(dataSeed[0][i]); j++)
            {
                data.Add(i+1);
            }
        }
        List<List<int>> permutations = GeneratePermutations(data);

        var targetProcess = new Dictionary<int , List<int>>();
        var targetGift = new Dictionary<int, List<int>>();
        var modCountDiv = data.Count;

        for (int i = 0; i < permutations.Count; i++)
        {
            var targetProcessTemp = new List<int>();
            var targetGiftTemp = new List<int>();
            for (int j = 0; j < 9*24; j++)
            {
                var modCount = (j+1) % modCountDiv;
                if (modCount == 0)
                {
                    modCount = modCountDiv;
                }
                if (j == 0)
                {
                    targetProcessTemp.Add(permutations[i][0]);
                    targetGiftTemp.Add(Convert.ToInt32(targetRankList[0][0]));
                }
                else
                {
                    var targetTemp = targetProcessTemp[j - 1] + permutations[i][modCount-1];
                    targetTemp %= 24;
                    if (targetTemp == 0)
                    {
                        targetTemp = 24;
                    }
                    targetProcessTemp.Add(targetTemp);
                    //获取价值量
                    var targetTemp2 = targetGiftTemp[j-1]+ Convert.ToInt32(targetRankList[targetTemp-1][0]);
                    targetGiftTemp.Add(targetTemp2);
                }
            }
            targetProcess[i]=targetProcessTemp;
            targetGift[i]=targetGiftTemp;
        }
        var filteredData = targetProcess;
        //过滤方案
        for (int i = 0; i < targetKeyList.Count; i++)
        {
            var rollTimes = targetKeyList[i][0];
            var rollGrid = targetKeyList[i][1];
            if (rollTimes != null)
            {
                var colIndex = Convert.ToInt32(rollTimes)-1;
                var colValue = Convert.ToInt32(rollGrid) ;
                //筛出指定列有目标值的行
                filteredData = filteredData
                    .Where(entry => entry.Value[colIndex] == colValue)
                    .ToDictionary(entry => entry.Key, entry => entry.Value);
                //去除非指定列有目标值的行
                var filterCondition = GenerateFilterConditions(colIndex, maxRoll ,colValue);
                filteredData = filteredData
                    .Where(entry => filterCondition.All(condition => condition(entry.Value)))
                    .ToDictionary(entry => entry.Key, entry => entry.Value);
            }
        }
        //方案整理
        var filteredDataGift = new List<List<object>>();
        var filteredDataMethod = new List<List<object>>();
        foreach (var key in filteredData.Keys)
        {
            filteredDataGift.Add(new List<object> { targetGift[key][maxRoll-1] });
            var methodStr = "";
            foreach (var method in permutations[key])
            {
                methodStr += method.ToString()+",";
            }
            methodStr.Substring(0, methodStr.Length - 1);
            filteredDataMethod.Add(new List<object> { methodStr });
        }
        PubMetToExcel.ListToArrayToRange(filteredDataGift, ws,11, 6);
        PubMetToExcel.ListToArrayToRange(filteredDataMethod, ws, 11, 5);

        // 释放 COM 对象
        Marshal.ReleaseComObject(ws);
        Marshal.ReleaseComObject(Wk);
        Marshal.ReleaseComObject(CreatRibbon._app);

        //Debug.Print("All Permutations:");
        //foreach (var permutation in permutations)
        //{
        //    Debug.Print(string.Join(", ", permutation));
        //}
    }
    static List<Func<List<int>, bool>> GenerateFilterConditions(int onlyCol,int otherCol,int conditionValue)
    {
        // 在实际情况下，你可以根据动态条件生成适当的委托列表
        var conditions = new List<Func<List<int>, bool>>();

        //conditions.Add(values => values[onlyCol] == conditionValue);
        //添加动态条件的示例
        for (int i = 0; i < otherCol; i++)
        {
            if (i != onlyCol)
            {
                conditions.Add(values => values[i] != conditionValue);
            }
        }
        return conditions;
    }

    static List<List<int>> GeneratePermutations(List<int> data)
    {
        List<List<int>> result = new List<List<int>>();
        GeneratePermutationsHelper(data, 0, result);
        return result;
    }

    static void GeneratePermutationsHelper(List<int> data, int index, List<List<int>> result)
    {
        if (index == data.Count)
        {
            // 当索引达到数组末尾时，添加当前排列到结果集
            result.Add(new List<int>(data));
            return;
        }
        // 使用 HashSet 来去重
        HashSet<int> usedValues = new HashSet<int>();

        for (int i = index; i < data.Count; i++)
        {
            if (usedValues.Add(data[i]))
            {
                // 交换当前位置和索引位置的元素
                Swap(data, index, i);

                // 递归生成下一个位置的排列
                GeneratePermutationsHelper(data, index + 1, result);
                // 恢复交换的元素，以便进行下一次交换
                Swap(data, index, i);
            }
        }
    }
    static void Swap(List<int> data, int i, int j)
    {
        int temp = data[i];
        data[i] = data[j];
        data[j] = temp;
    }

    public static void TestCAPI()
    {
        Wk.Cells[1,1].Value = "Hello World";



    }

}
