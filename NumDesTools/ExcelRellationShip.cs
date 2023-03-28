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
    static int dataCount = 1;
    public static void ExcelDic()
    {
        excelLinkDictionary = new Dictionary<string, List<string>>();
        excelFixKeyDictionary = new Dictionary<string, List<int>>();
        excelFixKeyMethodDictionary =new Dictionary<string, List<string>>();
        Worksheet sheet = IndexWk.ActiveSheet;
        //读取模板表数据
        var rowsCount = (sheet.Cells[sheet.Rows.Count, "A"].End[XlDirection.xlUp].Row - 15) / 3;
        for (int i = 1; i <= rowsCount; i++)
        {
            var baseExcel = sheet.Cells[1, 1].Offset[15 + (i - 1) * 4, 0].Value.ToString();
            excelLinkDictionary[baseExcel] = new List<string>();
            excelFixKeyDictionary[baseExcel] = new List<int>();
            excelFixKeyMethodDictionary[baseExcel] =new List<string>();
            for (int j = 1; j <= 2; j++)
            {
                var linkExcel = sheet.Cells[1, 1].Offset[16 + (i - 1) * 4, j + 1].Value;
                var baseExcelFixKey = sheet.Cells[1, 1].Offset[17 + (i - 1) * 4, j + 1].Value;
                var baseExcelFixKeyMethod = sheet.Cells[1, 1].Offset[18 + (i - 1) * 4, j + 1].Value;
                excelLinkDictionary[baseExcel].Add(linkExcel);
                excelFixKeyDictionary[baseExcel].Add(Convert.ToInt32(baseExcelFixKey));
                excelFixKeyMethodDictionary[baseExcel].Add(baseExcelFixKeyMethod);
            }
        }
    }

    public static void test()
    {
        //FixValueType();
        ExcelDic();
        List<List<(int, int)>> modeIDRow = new List<List<(int, int)>>();
        var abc = new List<(int,int)>();
        abc.Add((1,10));
        modeIDRow.Add(abc);
        List<string> fileName = new List<string>();
        fileName.Add("索引1.xlsx");
        List<List<(int, int)>> modeID222 = new List<List<(int, int)>>();
        List<(int, int)> modeID = new List<(int, int)>();
        modeID.Add((1,323));
        modeID222.Add(modeID);
        var sheet = IndexWk.ActiveSheet;
        string WriteMode = sheet.Range["B11"].value.ToString();
        string testKey = sheet.Range["C11"].value.ToString();
        CreateRellationShip(fileName, modeIDRow, WriteMode, testKey, modeID222);

        //test2(fileName);
        //var excel = new FileStream(excelPath + @"\索引1.xlsx", FileMode.Open, FileAccess.Read);
        //var workbook = new XSSFWorkbook(excel);
        //var sheet = workbook.GetSheetAt(0);
        //sheet.ShiftRows(1, sheet.LastRowNum, 1, true, false);
        //IRow row = sheet.CreateRow(1);
        //ICell cell1 = row.CreateCell(0);
        //cell1.SetCellValue("New Cell 1");
        //var excel2 = new FileStream(excelPath + @"\索引1.xlsx", FileMode.Create, FileAccess.Write);
        //workbook.Write(excel2);
        //workbook.Close();
        //excel2.Close();
        //excel.Close();
        //var asd =FindSourceRow(sheet, 1, "10");

        //var excel = new FileStream(excelPath + @"\" + "样表起始表.xlsx", FileMode.Open, FileAccess.Read);
        //var workbook = new XSSFWorkbook(excel);
        //var sheet = workbook.GetSheetAt(0);
        //for(int i =0;i<sheet.LastRowNum;i++)
        //{
        //    var row = sheet.GetRow(i);
        //    var cell = row.GetCell(0);
        //    if(cell == null) continue;
        //    Debug.Print(cell.ToString());
        //}

    }

    public static string ValueTypeToStringInNPOI(ICell cell)
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
                default:
                    cellValueAsString = "";
                    break;
            }
        }

        return cellValueAsString;
    }

    public static int FindSourceRow(ISheet sheet, int col, string searchValue)
    {
        for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);
            if (row != null)
            {
                var cell = row.GetCell(col);
                var cellValue = ValueTypeToStringInNPOI(cell);
                if (cellValue == searchValue)
                {
                    return i;
                }
            }
        }
        return -1;
    }
    public static string RegNumReplaceNew(string text, List<(int,int)> digit,bool isCarry,List<(int, int)> keyBitCount)
    {
        var numCount=1;
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
                var newNum = num + (int)Math.Pow(10, digit[ditCount].Item2 - 1);
                if (isCarry == false)
                {
                    int digitCount = (int)Math.Log10(newNum + 1) + 1 - keyBitCount[numCount-1].Item1; //数字位数-----需要更改i和原始数字相关，不能简单的只取一位了；；需要取得原来数字长度
                    int digitValue = num / (int)Math.Pow(10, digit[ditCount].Item2 - 1) % (int)Math.Pow(10, digitCount + 1); // 获取要增加的数字位的值
                    if (digitValue + 1 >= (int)Math.Pow(10, digitCount + 1))
                    {
                        var newnumber = num / (int)Math.Pow(10, digit[ditCount].Item2 + digitCount);
                        var mumberMOd = num % (int)Math.Pow(10, digit[ditCount].Item2 - 1);
                        var newdigitValue = (digitValue + 1) * (int)Math.Pow(10, digit[ditCount].Item2 - 1);
                        newnumber = newnumber * (int)Math.Pow(10, digit[ditCount].Item2 + digitCount + 1) + newdigitValue + mumberMOd;
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
            else if(digit.Count == 1 && digit[0].Item1 == 0)
            {
                var newNum = num + (int)Math.Pow(10, digit[0].Item2 - 1);
                text = text.Replace(numStr, newNum.ToString());
            }
            numCount++;
        }
        return text;
    }
    public static string RegNumReplaceNew2(string text, int digit)
    {

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
    public static void CreateRellationShip(List<string> oldFileName, List<List<(int, int)>> oldmodelID, string WriteMode, string testKey, List<List<(int, int)>> oldExcelIDGroup)
    {
        List<List<(int, int)>> newmodelID = new List<List<(int, int)>>();
        List<string> newFileName = new List<string>();
        List<List<(int, int)>> newExcelID = new List<List<(int, int)>>();
        int excount = 0;
        foreach (var excelFile in oldFileName)
        {
            var excel = new FileStream(excelPath + @"\" + excelFile, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(excel);
            var sheet = workbook.GetSheetAt(0);
            for (int k = 0; k < oldmodelID[excount].Count; k++)
            {
     
                    var seachValue = oldmodelID[excount][k].Item2;
                    var rowReSourceRow = FindSourceRow(sheet, 1, seachValue.ToString());
                    if (rowReSourceRow == -1) continue;
                    //数据复制
                    for (int i = 0; i < dataCount; i++)
                    {
                        var rowSource = sheet.GetRow(rowReSourceRow) ?? sheet.CreateRow(rowReSourceRow);
                        var colTotal = sheet.GetRow(1).LastCellNum + 1;
                        if (WriteMode == "新增")
                        {
                            if (sheet.LastRowNum != rowReSourceRow)
                            {
                                sheet.ShiftRows(rowReSourceRow + 1, sheet.LastRowNum, dataCount, true, false);
                            }
                        }

                        var rowTarget = sheet.GetRow(rowReSourceRow + i + 1) ??
                                            sheet.CreateRow(rowReSourceRow + i + 1);
                        var rowTargetTemp = sheet.GetRow(rowReSourceRow + i + 1) ??
                                                sheet.CreateRow(rowReSourceRow + i + 1);
                        for (int j = 0; j < colTotal; j++)
                        {
                            var cellSource = rowSource.GetCell(j) ?? rowSource.GetCell(j);
                            string cellSourceValue;
                            if (cellSource != null)
                            {
                                cellSourceValue = ValueTypeToStringInNPOI(cellSource);
                                var cellTarget = rowTarget.GetCell(j) ?? rowTarget.CreateCell(j);
                                //if(WriteMode=="修改") continue;
                                //表格的ID字段的修改--后续要添加其他字段的更改方式
                                if (j == 1)
                                {
                                    var asdasd = oldExcelIDGroup[excount][k].Item2;
                                    cellTarget.SetCellValue(asdasd);
                                }
                                else
                                {
                                    cellTarget.SetCellValue(cellSourceValue);
                                }

                                cellTarget.CellStyle = cellSource.CellStyle;
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
                                if (excelFileFixKey == 0) continue;
                                var cellTarget = sheet.GetRow(rowReSourceRow).GetCell(excelFileFixKey);
                                var cellTargetValue = ValueTypeToStringInNPOI(cellTarget);
                                //修改字段字典中的字段值，各自方法不一
                                var cellFixValueIdList = new List<(int, int)>();
                                var cellSourceValueList = new List<(int, int)>();
                                var rowFix = sheet.GetRow(rowReSourceRow + i + 1) ?? sheet.CreateRow(rowReSourceRow + i + 1);
                                var cellFix = rowFix.GetCell(excelFileFixKey) ?? rowFix.CreateCell(excelFileFixKey);
                                var cellFixValue = ValueTypeToStringInNPOI(cellFix);
                                //每个字段的Value修改方式不一，需要调用方法:检测string是否有[，如果有则需要正则把所有的数值提取出来并替换

                                //字段每个数字位数统计，原始modeID统计
                                cellSourceValueList = KeyBitCount(cellFixValue);
                                //字段值改写方法
                                var tmep2 = CellFixValueKeyList(excelFixKeyMethodDictionary[excelFile][indexExcelCount]);
                                //修改字符串
                                var cellFixValue2 =
                                    RegNumReplaceNew(cellFixValue, tmep2, false, cellSourceValueList);
                                //统计新ID
                                cellFixValueIdList = KeyBitCount(cellFixValue2);

                                cellFix.SetCellValue(cellFixValue2);
                                cellFix.CellStyle = cellFix.CellStyle;
                                //有关联表的字段的ID传递出去
                                //表格关联字典中寻找下一个递归文件，有关联表的字段ID要生成List递归
                                if (indexExcel != null)
                                {
                                    newFileName.Add(indexExcel);
                                    newmodelID.Add(cellSourceValueList);
                                    newExcelID.Add(cellFixValueIdList);
                                }
                                indexExcelCount++;
                            }
                        }
                    }
                
            }
            var excel2 = new FileStream(excelPath + @"\" + oldFileName[excount], FileMode.Create, FileAccess.Write);
            workbook.Write(excel2);
            workbook.Close();
            excel2.Close();
            excel.Close();
            excount++;
        }
        if (newFileName.Count > 0)
        {
            CreateRellationShip(newFileName, newmodelID, WriteMode, testKey, newExcelID);
        }
    }


    public static void FixValueType()
    {

        //string str = "1#2,3#2,2"; // 要处理的字符串
        string str = "1#2,3#2,4"; // 要处理的字符串
        var tempList = CellFixValueKeyList(str);

        var str1 = "[1001,1002,1003,1004]";

        var keyBitCount = KeyBitCount(str1);

        for (int i = 0; i < 20; i++)
        {
            str1 = RegNumReplaceNew(str1, tempList, false, keyBitCount);
            Debug.Print(str1);

        }


    }

    private static List<(int,int)> KeyBitCount(string str)
    {
        Regex regex = new Regex(@"\d+");
        var matches = regex.Matches(str);
        var keyBitCount = new List<(int,int)>();
        foreach (var matche in matches)
        {
            var temp = matche.ToString();
            int digitCount = (int)Math.Log10(Convert.ToInt32(temp) + 1) + 1;
            keyBitCount.Add((digitCount, Convert.ToInt32(temp)));
        }

        return keyBitCount;
    }

    private static List<(int,int)> CellFixValueKeyList(string str)
    {
        var numkeyList = new List<(int,int)>();

        string[] pairs;
        if (str==null)
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
                    numkeyList.Add((key,value));
                }
                else
                {
                    numkeyList.Add((int.Parse(pair), 1));
                }
            }
        }
        else
        {
            int strtemp;
            if (str == "")
            {
                strtemp = 0;
            }
            else
            {
                strtemp=int.Parse(str);
            }
            numkeyList.Add((strtemp, 1));
        }
        return numkeyList;
    }
}

