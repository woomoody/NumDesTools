using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.Streaming.Values;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using static System.IO.Path;
using static NPOI.HSSF.Util.HSSFColor;
using ICell = NPOI.SS.UserModel.ICell;

namespace NumDesTools;


class ExcelRellationShip
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic IndexWk = App.ActiveWorkbook;
    private static readonly dynamic excelPath = IndexWk.Path;
    public static Dictionary<string, List<string>> excelLinkDictionary;
    public static Dictionary<string, List<int>> excelFixKeyDictionary;
    public static Dictionary<string, List<string>> excelFixKeyMethodDictionary;
    static dynamic sheet = IndexWk.ActiveSheet;
    static int dataCount = Convert.ToInt32(sheet.Range["E3"].value);
    public static void StartExcelData()
    {
        ExcelDic();
        var startModeId = sheet.Range["D3"].value;
        var startModeIdFix = sheet.range["F3"].value;
        List<List<(long, long)>> modeIDRow = new List<List<(long, long)>>();
        var tempList = new List<(long, long)>();
        tempList.Add((1, Convert.ToInt64(startModeId)));
        modeIDRow.Add(tempList);
        List<List<List<(long, long)>>> excelIDGroupStart = new List<List<List<(long, long)>>>();
        List<List<(long, long)>> excelIDGroupStartTemp1 = new List<List<(long, long)>>();
        List<(long, long)> excelIDGroupStartTemp2;
        for (int i = 0; i < dataCount; i++)
        {
            //modeID原始位数
            var temp2 = KeyBitCount(startModeId.ToString());
            //字段值改写方法
            if (startModeIdFix == null) startModeIdFix = "";
            var temp1 = CellFixValueKeyList(startModeIdFix.ToString());
            //修改字符串
            var cellFixValue2 =
                RegNumReplaceNew(startModeId.ToString(), temp1, false, temp2, 1 + i);
            excelIDGroupStartTemp2 = KeyBitCount(cellFixValue2);

            excelIDGroupStartTemp1.Add(excelIDGroupStartTemp2);
        }
        excelIDGroupStart.Add(excelIDGroupStartTemp1);
        string WriteMode = sheet.Range["B3"].value.ToString();
        List<string> fileName = new List<string>();
        fileName.Add(sheet.Range["C3"].value.ToString());

        CreateRellationShip(fileName, modeIDRow, WriteMode, excelIDGroupStart);
    }
    public static void ExcelDic()
    {
        excelLinkDictionary = new Dictionary<string, List<string>>();
        excelFixKeyDictionary = new Dictionary<string, List<int>>();
        excelFixKeyMethodDictionary = new Dictionary<string, List<string>>();
        Worksheet sheet = IndexWk.ActiveSheet;
        //读取模板表数据
        var rowsCount = (sheet.Cells[sheet.Rows.Count, "B"].End[XlDirection.xlUp].Row - 4) / 4 + 1;
        for (int i = 1; i <= rowsCount; i++)
        {
            var baseExcel = sheet.Cells[1, 1].Offset[4 + (i - 1) * 4, 1].Value.ToString();
            excelLinkDictionary[baseExcel] = new List<string>();
            excelFixKeyDictionary[baseExcel] = new List<int>();
            excelFixKeyMethodDictionary[baseExcel] = new List<string>();
            for (int j = 2; j <= 14; j++)
            {
                var linkExcel = sheet.Cells[1, 1].Offset[5 + (i - 1) * 4, j + 1].Value;
                var baseExcelFixKey = sheet.Cells[1, 1].Offset[6 + (i - 1) * 4, j + 1].Value;
                var baseExcelFixKeyMethod = sheet.Cells[1, 1].Offset[7 + (i - 1) * 4, j + 1].Value;
                excelLinkDictionary[baseExcel].Add(linkExcel);
                if (baseExcelFixKey == null)
                {
                    baseExcelFixKey = 0;
                }
                else if (baseExcelFixKey.ToString() == "")
                {
                    baseExcelFixKey = 0;
                }

                excelFixKeyDictionary[baseExcel].Add(Convert.ToInt32(baseExcelFixKey));
                if (baseExcelFixKeyMethod == null)
                {
                    baseExcelFixKeyMethod = "";
                }

                excelFixKeyMethodDictionary[baseExcel].Add(Convert.ToString(baseExcelFixKeyMethod));
            }
        }
    }
    public static string ValueTypeToStringInNPOI(ICell cell, XSSFWorkbook workbook)
    {
        string cellValueAsString = string.Empty;
        if (cell != null)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    cellValueAsString = cell.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    cellValueAsString = cell.StringCellValue;
                    break;
                case CellType.Boolean:
                    cellValueAsString = cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    cellValueAsString = cell.ErrorCellValue.ToString();
                    break;
                case CellType.Formula:
                    // Create a formula evaluator
                    var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
                    // Evaluate the formula and get the resulting value
                    var value = evaluator.Evaluate(cell).NumberValue;
                    cellValueAsString = value.ToString();
                    break;
                default:
                    cellValueAsString = "";
                    break;
            }
        }

        return cellValueAsString;
    }

    public static int FindSourceRow(ISheet sheet, int col, string searchValue, XSSFWorkbook workbook)
    {
        for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);
            if (row != null)
            {
                var cell = row.GetCell(col);
                var cellValue = ValueTypeToStringInNPOI(cell, workbook);
                if (cellValue == searchValue)
                {
                    return i;
                }
            }
        }

        return -1;
    }

    public static string RegNumReplaceNew(string text, List<(int, int)> digit, bool isCarry,
        List<(long, long)> keyBitCount, int addValue)
    {
        var numCount = 1;
        var ditCount = 0;
        var pattern = "\\d+";
        // 使用正则表达式匹配数字
        var matches = Regex.Matches(text, pattern);
        foreach (System.Text.RegularExpressions.Match match in matches)
        {
            var numStr = match.Value;
            var num = int.Parse(numStr);

            if (digit.Any(item => item.Item1 == numCount))
            {
                var newNum = num + (int)Math.Pow(10, digit[ditCount].Item2 - 1) * addValue;
                if (isCarry == false)
                {
                    var digitCount =
                        (long)Math.Log10(newNum + 1) + 1 -
                        keyBitCount[numCount - 1].Item1; //数字位数-----需要更改i和原始数字相关，不能简单的只取一位了；；需要取得原来数字长度
                    int digitValue = num / (int)Math.Pow(10, digit[ditCount].Item2 - 1) %
                                     (int)Math.Pow(10, digitCount + 1); // 获取要增加的数字位的值
                    if (digitValue + addValue >= (int)Math.Pow(10, digitCount + 1))
                    {
                        var newnumber = num / (int)Math.Pow(10, digit[ditCount].Item2 + digitCount);
                        var mumberMOd = num % (int)Math.Pow(10, digit[ditCount].Item2 - 1);
                        var newdigitValue = (digitValue + addValue) * (int)Math.Pow(10, digit[ditCount].Item2 - 1);
                        newnumber = newnumber * (int)Math.Pow(10, digit[ditCount].Item2 + digitCount + 1) +
                                    newdigitValue + mumberMOd;
                        text = text.Replace(numStr, newnumber.ToString());
                    }
                    else
                    {
                        text = text.Replace(numStr, newNum.ToString());
                    }
                }
                else
                {
                    text = text.Replace(numStr, newNum.ToString());

                }

                ditCount++;
            }
            else if (digit.Count == 1 && digit[0].Item1 == 0)
            {
                var newNum = num + (int)Math.Pow(10, digit[ditCount].Item2 - 1) * addValue;
                text = text.Replace(numStr, newNum.ToString());
            }

            numCount++;
        }

        return text;
    }
    //public static void test2(List<string> oldstr)
    //{
    //    List<string> newstr = new List<string>();
    //    foreach (var str in oldstr)
    //    {
    //        if (str == null) continue;
    //        if (excelLinkDictionary.ContainsKey(str))
    //        {
    //            foreach (var indestr in excelLinkDictionary[str])
    //            {
    //                newstr.Add(indestr);
    //            }
    //        }
    //        Debug.Print(str + "\n" + "\t");

    //    }
    //    if (newstr.Count > 0)
    //    {
    //        test2(newstr);
    //    }

    //}
    public static void CreateRellationShip(List<string> oldFileName, List<List<(long, long)>> oldmodelID,
        string WriteMode, List<List<List<(long, long)>>> oldExcelIDGroup)
    {
        List<List<(long, long)>> newmodelID = new List<List<(long, long)>>();
        List<string> newFileName = new List<string>();
        List<List<List<(long, long)>>> newExcelID = new List<List<List<(long, long)>>>();
        int excount = 0;
        foreach (var excelFile in oldFileName)
        {
            var excel = new FileStream(excelPath + @"\" + excelFile, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(excel);
            var sheet = workbook.GetSheetAt(0);
            for (int k = 0; k < oldmodelID[excount].Count; k++)
            {
                var seachValue = oldmodelID[excount][k].Item2;
                var rowReSourceRow = FindSourceRow(sheet, 1, seachValue.ToString(), workbook);
                if (rowReSourceRow == -1) continue;
                var rowSource = sheet.GetRow(rowReSourceRow) ?? sheet.CreateRow(rowReSourceRow);
                var colTotal = sheet.GetRow(1).LastCellNum + 1;
                if (WriteMode == "新增")
                {
                    if (sheet.LastRowNum != rowReSourceRow)
                    {
                        sheet.ShiftRows(rowReSourceRow + 1, sheet.LastRowNum, dataCount, true, false);
                    }
                }
                //数据复制
                for (int i = 0; i < dataCount; i++)
                {
                    var rowTarget = sheet.GetRow(rowReSourceRow + i + 1) ??
                                    sheet.CreateRow(rowReSourceRow + i + 1);
                    for (int j = 0; j < colTotal; j++)
                    {
                        var cellSource = rowSource.GetCell(j) ?? rowSource.GetCell(j);
                        string cellSourceValue;
                        if (cellSource != null)
                        {
                            cellSourceValue = ValueTypeToStringInNPOI(cellSource, workbook);
                            var cellTarget = rowTarget.GetCell(j) ?? rowTarget.CreateCell(j);
                            //if(WriteMode=="修改") continue;
                            //表格的ID字段的修改--后续要添加其他字段的更改方式
                            if (j == 1)
                            {
                                var tempValue = oldExcelIDGroup[excount][i][k].Item2;
                                cellTarget.SetCellValue(tempValue);
                                //Debug.Print(cellTarget.ToString());
                            }
                            else
                            {
                                cellTarget.SetCellValue(cellSourceValue);
                            }

                            cellTarget.CellStyle = cellSource.CellStyle;
                        }
                    }
                }

                if (excelFile == null) continue;
                if (excelLinkDictionary.ContainsKey(excelFile))
                {
                    var indexExcelCount = 0;
                    foreach (var indexExcel in excelLinkDictionary[excelFile])
                    {
                        var excelFileFixKey = excelFixKeyDictionary[excelFile][indexExcelCount];
                        //字典会把空值当0用
                        if (excelFileFixKey == 0)
                        {
                            indexExcelCount++;
                            continue;
                        }

                        //修改字段字典中的字段值，各自方法不一
                        var cellFixValueIdList = new List<List<(long, long)>>();
                        var cellSourceValueList = new List<(long, long)>();
                        var newMode = new List<(long, long)>();
                        for (int i = 0; i < dataCount; i++)
                        {
                            var rowFix = sheet.GetRow(rowReSourceRow + i + 1) ??
                                         sheet.CreateRow(rowReSourceRow + i + 1);
                            var cellFix = rowFix.GetCell(excelFileFixKey) ?? rowFix.CreateCell(excelFileFixKey);
                            var cellFixValue = ValueTypeToStringInNPOI(cellFix, workbook);
                            //if (excelFile == "FlotageGroupData.xlsx")
                            //{
                            //    var abc = 1;
                            //}
                            //每个字段的Value修改方式不一，需要调用方法:检测string是否有[，如果有则需要正则把所有的数值提取出来并替换

                            //字段每个数字位数统计，原始modeID统计
                            cellSourceValueList = KeyBitCount(cellFixValue);
                            //字段值改写方法
                            var temp1 = CellFixValueKeyList(excelFixKeyMethodDictionary[excelFile][indexExcelCount]);
                            //修改字符串
                            var cellFixValue2 =
                                RegNumReplaceNew(cellFixValue, temp1, false, cellSourceValueList, 1 + i);
                            //统计新ID
                            var temp2 = KeyBitCount(cellFixValue2);
                            //标记重复项
                            newMode = cellSourceValueList.Except(temp2).ToList();
                            var newFix = temp2.Except(cellSourceValueList).ToList();

                            cellFixValueIdList.Add(newFix);
                            cellFix.SetCellValue(cellFixValue2);
                            //cellFix.CellStyle = cellFix.CellStyle;
                        }

                        //有关联表的字段的ID传递出去
                        //表格关联字典中寻找下一个递归文件，有关联表的字段ID要生成List递归
                        if (indexExcel != null && indexExcel != "")
                        {
                            newFileName.Add(indexExcel);
                            newmodelID.Add(newMode);
                            newExcelID.Add(cellFixValueIdList);
                        }

                        indexExcelCount++;
                    }
                }
            }

            //if (WriteMode == "新增")
            //{
            //    //去重
            //    // 假设要筛选的列为第一列（从0开始编号）
            //    var columnToFilter = 1;
            //    var rowIndexToStart = 4; // 第一行是表头，从第二行开始筛选

            //    var existingValues = new HashSet<string>(); // 用于记录已经出现的值

            //    for (int i = rowIndexToStart; i <= sheet.LastRowNum; i++)
            //    {
            //        var row = sheet.GetRow(i);
            //        if (row == null) continue;

            //        var cell = row.GetCell(columnToFilter);
            //        if (cell == null) continue;

            //        var cellValue = cell.ToString();
            //        if (existingValues.Contains(cellValue))
            //        {
            //            // 如果值已经出现过，则删除当前行

            //            sheet.ShiftRows(i + 1, sheet.LastRowNum, -1);

            //            i--; // 因为删除了一行，需要将计数器 i 减1
            //        }
            //        else
            //        {
            //            // 否则，将值加入集合
            //            existingValues.Add(cellValue);
            //        }
            //    }
            //}

            var excel2 = new FileStream(excelPath + @"\" + oldFileName[excount], FileMode.Create, FileAccess.Write);
            workbook.Write(excel2);
            workbook.Close();
            excel2.Close();
            excel.Close();
            excount++;
        }

        if (newFileName.Count > 0)
        {
            CreateRellationShip(newFileName, newmodelID, WriteMode, newExcelID);
        }
    }

    public static void FixValueType()
    {

        //string str = "1#2,3#2,2"; // 要处理的字符串
        string str = "1#2,3#2,4"; // 要处理的字符串
        var tempList = CellFixValueKeyList(str);

        var str1 = "[11001,11002,10003,10004]";

        //var keyBitCount = KeyBitCount(str1);

        //for (int i = 0; i < 20; i++)
        //{
        //    str1 = RegNumReplaceNew(str1, tempList, false, keyBitCount, 2);
        //    Debug.Print(str1);

        //}


    }

    private static List<(long, long)> KeyBitCount(string str)
    {
        Regex regex = new Regex(@"\d+");
        var matches = regex.Matches(str);
        var keyBitCount = new List<(long digitCount, long)>();
        foreach (var matche in matches)
        {
            var temp = matche.ToString();
            var digitCount = (long)Math.Log10(Convert.ToInt64(temp) + 1) + 1;
            keyBitCount.Add((digitCount, Convert.ToInt64(temp)));
        }

        return keyBitCount;
    }

    private static List<(int, int)> CellFixValueKeyList(string str)
    {
        var numkeyList = new List<(int, int)>();

        string[] pairs;
        if (str == null)
        {
            str = "";
        }

        if (str.Contains(','))
        {
            pairs = str.Split(','); // 将字符串按逗号分隔成多个键值对
            foreach (string pair in pairs)
            {
                string[] parts;
                if (pair.Contains('#'))
                {
                    parts = pair.Split('#'); // 将键值对按井号分隔成键和值
                    int key;
                    if (!int.TryParse(parts[0], out key)) // 尝试将值解析为整数，如果解析失败就将值设为 0
                    {
                        MessageBox.Show(str + "#前必须有数值");
                        Environment.Exit(0);
                    }

                    int value;
                    if (!int.TryParse(parts[1], out value)) // 尝试将值解析为整数，如果解析失败就将值设为 0
                    {
                        value = 1;
                    }

                    numkeyList.Add((key, value));
                }
                else
                {
                    numkeyList.Add((int.Parse(pair), 1));
                }
            }
        }
        else
        {
            if (str.Contains('#'))
            {
                int key;
                int value;
                var parts = str.Split('#');
                key = Convert.ToInt32(parts[0]);
                value = Convert.ToInt32(parts[1]);
                numkeyList.Add((key, value));
            }
            else
            {
                int strtemp;
                if (str == "")
                {
                    strtemp = 0;
                    numkeyList.Add((strtemp, 1));
                }
                else
                {
                    strtemp = int.Parse(str);
                    numkeyList.Add((0, strtemp));
                }
            }
        }

        return numkeyList;
    }

    public static void ExcelHyperLinks()
    {
        var sheet = IndexWk.ActiveSheet;
        for (int i = 3; i <= 101; i++)
        {
            for (int j = 2; j <= 20; j++)
            {
                var cell = sheet.Cells[i, j];
                var cs1 = cell.Style;
                if (cell.value != null && cell.value.ToString().Contains(".xlsx"))
                {
                    cell.Hyperlinks.Add(cell, excelPath + @"\" + cell.value.ToString());
                    cell.Font.Size = 9;
                    cell.Font.Name = "微软雅黑";
                    //cell.Copy();
                    //cell.PasteSpecial(XlPasteType.xlPasteFormats);
                    //cell.Style = cs1;
                }
            }
        }
    }
}

