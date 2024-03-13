using ExcelDna.Integration;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.Office.Core;
using System;
using System.Diagnostics;
using System.Windows;
using Microsoft.VisualStudio.TextManager.Interop;
using OfficeOpenXml;

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类调用的具体业务逻辑
/// </summary>
public class PubMetToExcelFunc
{
    private static readonly dynamic Wk = NumDesAddIn.App.ActiveWorkbook;

    private static readonly string Path = Wk.Path;

    //Excel数据查询并合并表格数据
    public static void ExcelDataSearchAndMerge(string searchValue)
    {
        //获取所有的表格路径
        string[] ignoreFileNames = ["#", "副本"];
        var rootPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(Path));
        var fileList = new List<string>()
            { rootPath + @"\Excels\Tables\", rootPath + @"\Excels\Localizations\", rootPath + @"\Excels\UIs\" };
        var files = PubMetToExcel.PathExcelFileCollect(fileList, "*.xlsx", ignoreFileNames);
        //查找指定关键词，记录行号和表格索引号
        var findValueList = new List<(string, string, int, int, string, string)>();
        Parallel.ForEach(files, file =>
            {
                var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(file);
                var findValue = PubMetToExcel.FindDataInDataTable(file, dataTable, searchValue);
                if (findValue.Count > 0)
                    //findValueList.Add(findValue);
                    findValueList = findValueList.Concat(findValue).ToList();
            }
        );
        //人工查询所需要的数据，可以打开表格，可以删除和手动增加数据，专用表格进行操作
        dynamic tempWorkbook;
        try
        {
            tempWorkbook = NumDesAddIn.App.Workbooks.Open(rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
        }
        catch
        {
            tempWorkbook = NumDesAddIn.App.Workbooks.Add();
            tempWorkbook.SaveAs(rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
        }

        var tempSheet = tempWorkbook.Sheets["Sheet1"];
        var tempDataArray = new string[findValueList.Count, 5];
        for (var i = 0; i < findValueList.Count; i++)
        {
            tempDataArray[i, 0] = findValueList[i].Item1;
            tempDataArray[i, 1] = findValueList[i].Item2;
            tempDataArray[i, 2] = PubMetToExcel.ConvertToExcelColumn(findValueList[i].Item4) + findValueList[i].Item3;
            tempDataArray[i, 3] = findValueList[i].Item5;
            tempDataArray[i, 4] = findValueList[i].Item6;
        }

        var tempDataRange = tempSheet.Range[tempSheet.Cells[2, 2],
            tempSheet.Cells[2 + tempDataArray.GetLength(0) - 1, 2 + tempDataArray.GetLength(1) - 1]];
        tempDataRange.Value = tempDataArray;
        tempWorkbook.Save();
        //合并数据
    }

    //Excel右键识别文件路径并打开
    public static void RightOpenExcelByActiveCell(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sheet = NumDesAddIn.App.ActiveSheet;
        var selectCell = NumDesAddIn.App.ActiveCell;
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var workBookName = workBook.Name;
        var workbookPath = workBook.Path;
        workbookPath = System.IO.Path.GetDirectoryName(workbookPath);
        var selectCellValue = "";
        if (selectCell.Value != null) selectCellValue = selectCell.Value.ToString();
        //正则出是Excel路径的单元格
        var isMatch = selectCellValue.Contains(".xls");
        if (isMatch)
        {
            string sheetName;
            var cellAddress = "A1";
            if (workBookName.Contains("#合并表格数据缓存") )
            {
                var selectRow = selectCell.Row;
                var selectCol = selectCell.Column;
                sheetName = sheet.Cells[selectRow, selectCol + 1].Value;
                cellAddress = sheet.Cells[selectRow, selectCol + 2].Value;
            }
            else if(selectCellValue.Contains("#"))
            {
                var excelSplit = selectCellValue.Split("#");
                selectCellValue = workbookPath + @"\Tables\" + excelSplit[0];
                sheetName = excelSplit[1];
            }
            else
            {
                switch (selectCellValue)
                {
                    case "Localizations.xlsx":
                        selectCellValue = workbookPath + @"\Localizations\Localizations.xlsx";
                        break;
                    case "UIConfigs.xlsx":
                        selectCellValue = workbookPath + @"\UIs\UIConfigs.xlsx";
                        break;
                    case "UIItemConfigs.xlsx":
                        selectCellValue = workbookPath + @"\UIs\UIItemConfigs.xlsx";
                        break;
                    default:
                        selectCellValue = workbookPath + @"\Tables\" + selectCellValue;
                        break;
                }
                sheetName = "Sheet1";
            }
            PubMetToExcel.OpenExcelAndSelectCell(selectCellValue, sheetName, cellAddress);
        }
    }

    public static void OpenBaseLanExcel(CommandBarButton ctrl, ref bool cancelDefault)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var selectCell = NumDesAddIn.App.ActiveCell;
        var basePath = NumDesAddIn.App.ActiveWorkbook.Path;
        var newPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(basePath));
        newPath = newPath + @"\Excels\Localizations\Localizations.xlsx";
        var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(newPath);
        var findValue = PubMetToExcel.FindDataInDataTable(newPath, dataTable, selectCell.Value.ToString());
        var cellAddress = PubMetToExcel.ConvertToExcelColumn(findValue[0].Item4) + findValue[0].Item3;
        PubMetToExcel.OpenExcelAndSelectCell(newPath, "Sheet1", cellAddress);
    }

    public static void OpenMergeLanExcel(CommandBarButton ctrl, ref bool cancelDefault)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var selectCell = NumDesAddIn.App.ActiveCell;
        var basePath = NumDesAddIn.App.ActiveWorkbook.Path;
        var mergePath = "";
        //数据源路径txt
        var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePath = System.IO.Path.Combine(documentsFolder, "mergePath.txt");
        var mergePathList = PubMetToExcel.ReadWriteTxt(filePath);
        //第一行Alice，第二行Cove
        if (mergePathList.Count <= 1)
            //打开文本文件
            Process.Start(filePath);
        if (mergePathList[0] == "" || mergePathList[1] == "" || mergePathList[1] == mergePathList[0])
            //打开文本文件
            Process.Start(filePath);
        else
            mergePath = basePath != mergePathList[1] ? mergePathList[1] : mergePathList[0];
        var newPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(mergePath));
        newPath = newPath + @"\Localizations\Localizations.xlsx";
        var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(newPath);
        var findValue = PubMetToExcel.FindDataInDataTable(newPath, dataTable, selectCell.Value.ToString());
        string cellAddress;
        if (findValue.Count == 0)
            cellAddress = "A1";
        else
            cellAddress = PubMetToExcel.ConvertToExcelColumn(findValue[0].Item4) + findValue[0].Item3;
        PubMetToExcel.OpenExcelAndSelectCell(newPath, "Sheet1", cellAddress);
    }

    public static void AliceBigRicherDfs2(string sheetName)
    {

        var baseName = "大富翁种";
        if (!sheetName.Contains(baseName))
        {
            MessageBox.Show("当前表格不是【大富翁种**】,无法使用大富翁功能");
        }
        //读取数据（0起始）
        var targetRank = PubMetToExcel.ReadExcelDataC(sheetName, 21, 44, 2, 2);
        //object[,] seedRangeValue = PubMetToExcel.ReadExcelDataC(sheetName, 2, 7, 6, 6);
        var targetKey = PubMetToExcel.ReadExcelDataC(sheetName, 13, 15, 2, 3);
        var targetKeySoft = PubMetToExcel.ReadExcelDataC(sheetName, 2, 11, 2, 3);
        var maxRollCell = PubMetToExcel.ReadExcelDataC(sheetName, 18, 18, 2, 2);
        var maxGridLoopCell = PubMetToExcel.ReadExcelDataC(sheetName, 17, 17, 2, 2);
        var maxRankCell = PubMetToExcel.ReadExcelDataC(sheetName, 18, 18, 5, 5);
        var maxRoll = Convert.ToInt32(maxRollCell[0, 0]);
        var maxGridLoop = Convert.ToInt32(maxGridLoopCell[0, 0]);
        var maxRankValue = Convert.ToInt32(maxRankCell[0, 0]);
        //List<int> data = new List<int>();
        //for (int i = 0; i < seedRangeValue.GetLength(0); i++)
        //{
        //    var seed = seedRangeValue[i,0];
        //    for (int j = 0; j < Convert.ToInt32(seed); j++)
        //    {
        //        data.Add(i + 1);
        //    }
        //}
        var permutations = GenerateUniqueSchemes(maxRoll, maxRoll * 100000);
        var targetProcess = new Dictionary<int, List<int>>();
        var bpProcess = new Dictionary<int, List<int>>();
        var targetGift = new Dictionary<int, List<int>>();
        var modCountDiv = maxRoll;

        for (var i = 0; i < permutations.Count; i++)
        {
            var targetProcessTemp = new List<int>();
            var bpProcessTemp = new List<int>();
            var targetGiftTemp = new List<int>();
            //生成多少个格子
            for (var j = 0; j < maxGridLoop * 24; j++)
            {
                var modCount = (j + 1) % modCountDiv;
                if (modCount == 0) modCount = modCountDiv;
                if (j == 0)
                {
                    targetProcessTemp.Add(permutations[i][0]);
                    bpProcessTemp.Add(permutations[i][0]);
                    var tempValue = targetRank[0, 0];
                    if (tempValue is ExcelEmpty) tempValue = null;
                    targetGiftTemp.Add(Convert.ToInt32(tempValue));
                }
                else
                {
                    var targetTemp = targetProcessTemp[j - 1] + permutations[i][modCount - 1];
                    var processTemp = bpProcessTemp[j - 1] + permutations[i][modCount - 1];
                    bpProcessTemp.Add(processTemp);
                    targetTemp %= 24;
                    if (targetTemp == 0) targetTemp = 24;
                    targetProcessTemp.Add(targetTemp);
                    //获取价值量
                    var tempValue = targetRank[targetTemp - 1, 0];
                    if (tempValue is ExcelEmpty) tempValue = null;
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
        for (var i = 0; i < targetKey.GetLength(0); i++)
        {
            var rollTimes = targetKey[i, 0];
            var rollGrid = targetKey[i, 1];
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
                //filteredData = filteredData
                //    .Where(pair => pair.Value.Take(maxRoll).Count(item => item == colValue) == 1)
                //    .ToDictionary(pair => pair.Key, pair => pair.Value);
            }
        }

        //过滤动态目标
        for (var i = 0; i < targetKeySoft.GetLength(0); i++)
        {
            var softTimes = targetKeySoft[i, 1];
            var softGrid = targetKeySoft[i, 0];
            if (!(softGrid is ExcelEmpty))
                //筛选动态目标值满足出现次数的行
                filteredData = filteredData
                    .Where(pair =>
                        pair.Value.Take(maxRoll).Count(item => item == Convert.ToInt32(softGrid)) ==
                        Convert.ToInt32(softTimes))
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
        }

        //方案整理
        var filteredDataGift = new Dictionary<int, int>();
        var filteredDataMethod = new List<List<object>>();
        var filteredDataBpProcess = new List<List<object>>();
        foreach (var key in filteredData.Keys) filteredDataGift[key] = targetGift[key][maxRoll];
        //选择升阶进度中众数项
        //var modeValue = GetMode(filteredDataGift.Values);
        //选择升阶进度中指定值
        var modeValue = maxRankValue;
        var filteredDataGiftMode = filteredDataGift.Where(pair => pair.Value == modeValue).ToList();
        var filteredDataGiftList = new List<List<object>>();
        foreach (var kvp in filteredDataGiftMode)
        {
            var key = kvp.Key;
            var value = kvp.Value;
            filteredDataGiftList.Add([value]);
            filteredDataBpProcess.Add([bpProcess[key][maxRoll]]);
            var methodStr = "";
            foreach (var method in permutations[key]) methodStr += method + ",";
            methodStr = methodStr.Substring(0, methodStr.Length - 1);
            filteredDataMethod.Add([methodStr]);
        }

        //清理
        var emptyData = new object[65535 - 17 + 1, 6 - 6 + 1];
        PubMetToExcel.WriteExcelDataC(sheetName, 21, 65534, 4, 4, emptyData);
        PubMetToExcel.WriteExcelDataC(sheetName, 21, 65534, 5, 5, emptyData);
        PubMetToExcel.WriteExcelDataC(sheetName, 21, 65534, 6, 6, emptyData);
        //错误提示
        if (filteredDataBpProcess.Count == 0)
        {
            var error = new object[1, 1];
            error[0, 0] = "#Error#";
            PubMetToExcel.WriteExcelDataC(sheetName, 21, 21, 4, 4, error);
        }
        else
        {
            //写入
            PubMetToExcel.WriteExcelDataC(sheetName, 21, 21 + filteredDataBpProcess.Count - 1, 6, 6,
                PubMetToExcel.ConvertListToArray(filteredDataBpProcess));
            PubMetToExcel.WriteExcelDataC(sheetName, 21, 21 + filteredDataGiftList.Count - 1, 5, 5,
                PubMetToExcel.ConvertListToArray(filteredDataGiftList));
            PubMetToExcel.WriteExcelDataC(sheetName, 21, 21 + filteredDataMethod.Count - 1, 4, 4,
                PubMetToExcel.ConvertListToArray(filteredDataMethod));
        }
    }
    private static List<List<int>> GenerateUniqueSchemes(int numberOfRolls, int numberOfSchemes)
    {
        var result = new List<List<int>>();
        var seenSchemes = new HashSet<string>();
        var random = new Random();

        for (var i = 0; i < numberOfSchemes; i++)
        {
            var scheme = new List<int>();

            // 随机生成一个方案
            for (var j = 0; j < numberOfRolls; j++)
            {
                var randomNumber = random.Next(1, 7);
                scheme.Add(randomNumber);
            }

            // 转换为字符串，检查是否已经存在该方案
            var schemeString = string.Join(",", scheme);
            if (seenSchemes.Add(schemeString))
            {
                result.Add([..scheme]);
            }
        }
        return result;
    }
    //移动魔瓶模拟消耗
    public static void MagicBottleCostSimulate(string sheetName)
    {

        var baseName = "移动魔瓶";
        if (!sheetName.Contains(baseName))
        {
            MessageBox.Show("当前表格不是【移动魔瓶**】,无法使用魔瓶验算");
        }
        //读取数据（0起始）
        var eleCount = PubMetToExcel.ReadExcelDataC(sheetName, 2, 8, 21, 21);
        var simulateCount = PubMetToExcel.ReadExcelDataC(sheetName, 0, 0, 21, 21);
        var simulateCountMax = Convert.ToInt32(simulateCount[0, 0]);
        var eleCountMax = eleCount.GetLength(0);
        var filterEleCountMax = new List<int>();
        //初始化统计list
        for (var r = 0; r < eleCountMax; r++)
        {
            filterEleCountMax.Add(0);
        }
        //模拟猜数字
        for (var s = 0; s < simulateCountMax; s++)
        {
            for (var r = 0; r < eleCountMax; r++)
            {
                var eleGuessListGroup = new List<List<int>>();
                var eleNum = Convert.ToInt32(eleCount[r,0]);
                //创建随机元素序列List
                var eleList = new List<int>();
                var eleGuessList = new List<int>();
                for (int e = 1; e <= eleNum; e++)
                {
                    eleList.Add(e);
                    eleGuessList.Add(e);
                }
                var seedTarget = new Random();
                eleList = eleList.OrderBy(x => seedTarget.Next()).ToList();
                var seedGuess = new Random();
                eleGuessList = eleGuessList.OrderBy(x => seedGuess.Next()).ToList();
                do
                {
                    //随机猜数字，剔除对的
                    for (var eleCurrent = eleList.Count -1; eleCurrent >= 0; eleCurrent--)
                    {
                        var ele = eleList[eleCurrent];
                        var eleGuess = eleGuessList[eleCurrent];
                        if (eleGuess == ele)
                        {
                            eleList.RemoveAt(eleCurrent);
                            eleGuessList.RemoveAt(eleCurrent);
                        }
                        filterEleCountMax[r]++;
                    }
                    eleGuessListGroup.Add(eleGuessList);
                    //List重新排序和上次不同
                    if (eleList.Count > 1)
                    {
                        var eleTempList = new List<int>();
                        var seedTemp = new Random();
                        do
                        {
                            eleTempList = eleGuessList.OrderBy(x => seedTemp.Next()).ToList();
                        } while (eleGuessListGroup.Any(list => list.SequenceEqual(eleTempList)));
                        eleGuessList = eleTempList;
                    }
                } while (eleList.Count != 0);
            }
        }
        // ReSharper disable once PossibleLossOfFraction
        var filterEleCountMaxObj = filterEleCountMax.Select(item => (double)(item / simulateCountMax))
            .Select(simulateValue => new List<object> { simulateValue }).ToList();
        //清理
        var emptyData = new object[7, 7];
        PubMetToExcel.WriteExcelDataC(sheetName, 2, 8, 22, 22, emptyData);
        var emptyData2 = new object[1, 1];
        PubMetToExcel.WriteExcelDataC(sheetName, 0, 0, 23, 23, emptyData2);
        //错误提示
        if (filterEleCountMax.Count == 0)
        {
            var error = new object[1, 1];
            error[0, 0] = "#Error#";
            PubMetToExcel.WriteExcelDataC(sheetName, 0, 0, 23, 23, error);
        }
        else
        {
            //写入
            PubMetToExcel.WriteExcelDataC(sheetName, 2, 2 + filterEleCountMax.Count - 1, 22, 22,
                PubMetToExcel.ConvertListToArray(filterEleCountMaxObj));
        }
    }
    public static List<(string, string, string)> texstEncapsulation()
    {
        var errorList = new List<(string, string, string)>();
        string excelPath = @"C:\Users\cent\Desktop";
        string excelName = "tee.xlsx#Sheet1";
        var list = new List<dynamic>();
        //声明新类
        var excelObj = new ExcelDataByEpplus();
        //计算类属性
        excelObj.GetExcelObj(excelPath, excelName);
        //应用属性
        if (excelObj.ErrorList.Count > 0)
        {
            return excelObj.ErrorList;
        }
        //读取数据
        var sheet = excelObj.Sheet;
        List<dynamic> data = excelObj.Read(sheet, 5, 899);
        //修改数据
        for (int i = 0; i < data.Count; i++)
        {
            var firstRecord = (IDictionary<string, object>)data[i];
            foreach (var key in firstRecord.Keys.ToList())
            {
                // 尝试将值转换为字符串
                string stringValue = firstRecord[key]?.ToString();
                if (!string.IsNullOrEmpty(stringValue))
                {
                    // 如果字符串包含 "3"，则执行替换操作
                    firstRecord[key] = stringValue.Replace("20", "11");
                }
            }
        }
        //写入数据-如果要写到别的Excel则需要在声明新类
        var excel = excelObj.Excel;
        excelObj.Write(sheet, excel, data, 5);
        return errorList;
    }
}

