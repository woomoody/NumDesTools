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
using System.Windows;
using NPOI.SS.Util;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace NumDesTools;
/// <summary>
/// 公共的Excel功能类调用的具体业务逻辑
/// </summary>
public class PubMetToExcelFunc
{
    private static readonly dynamic Wk = CreatRibbon._app.ActiveWorkbook;
    private static readonly string path = Wk.Path;
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
    public static void AliceBigRicherDFS1()
    {
        var ws = Wk.ActiveSheet;
        object[,] targetRank  = ws.Range["D11:D34"].Value;
        var targetRankList = PubMetToExcel.RangeDataToList(targetRank);
        object[,] seedRangeValue = ws.Range["D2:I2"].Value;
        var dataSeed = PubMetToExcel.RangeDataToList(seedRangeValue);
        object[,] targetKey = ws.Range["D4:E8"].Value;
        var targetKeyList = PubMetToExcel.RangeDataToList(targetKey);
        int maxRoll = Convert.ToInt32(ws.Range["D9"].Value);
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
    public static void AliceBigRicherDFS2()
    {
        var sheetName = "Alice大富翁";
        //读取数据（0起始）
        object[,] targetRank = PubMetToExcel.ReadExcelDataC(sheetName, 16, 39, 2, 2);
        object[,] seedRangeValue = PubMetToExcel.ReadExcelDataC(sheetName, 2, 7, 6, 6);
        object[,] targetKey = PubMetToExcel.ReadExcelDataC(sheetName, 8, 10, 2, 3);
        object[,] targetKeySoft = PubMetToExcel.ReadExcelDataC(sheetName, 2, 6, 2, 3);
        object[,] maxRollCell = PubMetToExcel.ReadExcelDataC(sheetName, 13, 13, 2, 2);
        int maxRoll = Convert.ToInt32(maxRollCell[0, 0]);

        List<int> data = new List<int>();
        for (int i = 0; i < seedRangeValue.GetLength(0); i++)
        {
            var seed = seedRangeValue[i,0];
            for (int j = 0; j < Convert.ToInt32(seed); j++)
            {
                data.Add(i + 1);
            }
        }
        List<List<int>> permutations = GeneratePermutations(data);
        var targetProcess = new Dictionary<int, List<int>>();
        var bpProcess = new Dictionary<int, List<int>>();
        var targetGift = new Dictionary<int, List<int>>();
        var modCountDiv = data.Count;

        for (int i = 0; i < permutations.Count; i++)
        {
            var targetProcessTemp = new List<int>();
            var bpProcessTemp = new List<int>();
            var targetGiftTemp = new List<int>();
            //需要获取循环种子和24格子之间最小公倍数
            for (int j = 0; j < 9 * 24; j++)
            {
                var modCount = (j + 1) % modCountDiv;
                if (modCount == 0)
                {
                    modCount = modCountDiv;
                }
                if (j == 0)
                {
                    targetProcessTemp.Add(permutations[i][0]);
                    bpProcessTemp.Add(permutations[i][0]);
                    var tempValue = targetRank[0, 0];
                    if (tempValue is ExcelEmpty)
                    {
                        tempValue = null;
                    }
                    targetGiftTemp.Add(Convert.ToInt32(tempValue));
                }
                else
                {
                    var targetTemp = targetProcessTemp[j - 1] + permutations[i][modCount - 1];
                    var processTemp = bpProcessTemp[j - 1] + permutations[i][modCount - 1];
                    bpProcessTemp.Add(processTemp);
                    targetTemp %= 24;
                    if (targetTemp == 0)
                    {
                        targetTemp = 24;
                    }
                    targetProcessTemp.Add(targetTemp);
                    //获取价值量
                    var tempValue = targetRank[targetTemp - 1, 0];
                    if (tempValue is ExcelEmpty)
                    {
                        tempValue = null;
                    }
                    var targetTemp2 = targetGiftTemp[j - 1] + Convert.ToInt32(tempValue);
                    targetGiftTemp.Add(targetTemp2);
                }
            }
            targetProcess[i] = targetProcessTemp;
            bpProcess[i] = bpProcessTemp;
            targetGift[i] = targetGiftTemp;
        }
        var filteredData = targetProcess;
        //过滤固定目标
        for (int i = 0; i < targetKey.GetLength(0); i++)
        {
            var rollTimes = targetKey[i,0];
            var rollGrid = targetKey[i,1];
            if (!(rollTimes is ExcelEmpty))
            {
                var colIndex = Convert.ToInt32(rollTimes) - 1;
                var colValue = Convert.ToInt32(rollGrid);
                //筛选指定列有固定目标值的行
                filteredData = filteredData
                    .Where(entry => entry.Value[colIndex] == colValue)
                    .ToDictionary(entry => entry.Key, entry => entry.Value);
                //去除非指定列有固定目标值的行(换种思路，前XRoll只存在一个固定目标)
                                //var filterCondition = GenerateFilterConditions(colIndex, maxRoll, colValue);
                                //filteredData = filteredData
                                //    .Where(entry => filterCondition.All(condition => condition(entry.Value)))
                                //    .ToDictionary(entry => entry.Key, entry => entry.Value);
                filteredData = filteredData
                    .Where(pair => pair.Value.Take(maxRoll - 1).Count(item => item == colValue) == 1)
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
            }
        }
        //过滤动态目标
        for (int i = 0; i < targetKeySoft.GetLength(0); i++)
        {
            var softTimes = targetKeySoft[i,1];
            var softGrid = targetKeySoft[i,0];
            if (!(softGrid is ExcelEmpty))
            {
                //筛选动态目标值满足出现次数的行
                filteredData = filteredData
                    .Where(pair => pair.Value.Take(maxRoll - 1).Count(item => item == Convert.ToInt32(softGrid)) == Convert.ToInt32(softTimes))
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
            }
        }
        //方案整理
        var filteredDataGift = new Dictionary<int, int>();
        var filteredDataMethod = new List<List<object>>();
        var filteredDataBpProcess = new List<List<object>>();
        foreach (var key in filteredData.Keys)
        {
            filteredDataGift[key] = targetGift[key][maxRoll - 1];
        }
        //选择升阶进度中众数项
        int modeValue = GetMode(filteredDataGift.Values);
        var filteredDataGiftMode = filteredDataGift.Where(pair => pair.Value == modeValue).ToList();
        var filteredDataGiftList= new List<List<object>>();
        int methodCount = filteredDataGiftMode.Count;
        foreach (var kvp in filteredDataGiftMode)
        {
            int key = kvp.Key;
            int value = kvp.Value;
            filteredDataGiftList.Add(new List<object>{value});
            filteredDataBpProcess.Add(new List<object> { bpProcess[key][maxRoll - 1] });
            var methodStr = "";
            foreach (var method in permutations[key])
            {
                methodStr += method + ",";
            }
            methodStr = methodStr.Substring(0, methodStr.Length - 1);
            filteredDataMethod.Add(new List<object> { methodStr });
        }
        //清理
        object[,] emptyData = new object[65535-17+1, 6 - 6 +1];
        PubMetToExcel.WriteExcelDataC(sheetName, 16, 65534, 4, 4, emptyData);
        PubMetToExcel.WriteExcelDataC(sheetName, 16, 65534, 5, 5, emptyData);
        PubMetToExcel.WriteExcelDataC(sheetName, 16, 65534, 6, 6, emptyData);
        //错误提示
        if (filteredDataBpProcess.Count == 0)
        {
            var error = new object[1,1];
            error[0,0] = "#Error#";
            PubMetToExcel.WriteExcelDataC(sheetName,16,16,4,4, error);
        }
        else
        {
            //写入
            PubMetToExcel.WriteExcelDataC(sheetName, 16, 16 + filteredDataBpProcess.Count - 1, 6, 6, PubMetToExcel.ConvertListToArray(filteredDataBpProcess));
            PubMetToExcel.WriteExcelDataC(sheetName, 16, 16 + filteredDataGiftList.Count - 1, 5, 5, PubMetToExcel.ConvertListToArray(filteredDataGiftList));
            PubMetToExcel.WriteExcelDataC(sheetName, 16, 16 + filteredDataMethod.Count - 1, 4, 4, PubMetToExcel.ConvertListToArray(filteredDataMethod));
        }
    }
    public static void DataC()
    {
        object[,] result = new object[2,2];
        // 给数组赋予指定的值
        result[0, 0] = "A1";
        result[0, 1] = "B1";
        result[1, 0] = "A2";
        result[1, 1] = "B2";

        //读取数据（0起始）
        object[,] abc = PubMetToExcel.ReadExcelDataC("Sheet2", 0, 1, 0, 1);
        //写入数据（0起始）
        PubMetToExcel.WriteExcelDataC("Sheet2", 0, 1, 0, 1,result);
    }
    // 获取列表的众数
    static int GetMode(IEnumerable<int> values)
    {
        var modes = values
            .GroupBy(v => v)
            .OrderByDescending(g => g.Count())
            .Select(g => g.Key);

        // 如果有多个众数，可以在这里选择处理方式，此处选择返回第一个众数
        return modes.FirstOrDefault();
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
