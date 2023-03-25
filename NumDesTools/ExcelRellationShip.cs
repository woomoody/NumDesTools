using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.Common;
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
using NPOI.XWPF.UserModel;
using static System.IO.Path;
using ICell = NPOI.SS.UserModel.ICell;

namespace NumDesTools;


class ExcelRellationShip
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic IndexWk = App.ActiveWorkbook;
    private static readonly dynamic excelPath = IndexWk.Path;
    public static Dictionary<string, List<string>> excelLinkDictionary;
    public static Dictionary<string, List<int>> excelFixKeyDictionary;
    static int dataCount = 1;
    public static void ExcelDic()
    {
        excelLinkDictionary = new Dictionary<string, List<string>>();
        excelFixKeyDictionary = new Dictionary<string, List<int>>();
        Worksheet sheet = IndexWk.ActiveSheet;
        //读取模板表数据
        var rowsCount = (sheet.Cells[sheet.Rows.Count, "A"].End[XlDirection.xlUp].Row - 15) / 3;
        for (int i = 1; i <= rowsCount; i++)
        {
            var baseExcel = sheet.Cells[1, 1].Offset[15 + (i - 1) * 4, 0].Value.ToString();
            excelLinkDictionary[baseExcel] = new List<string>();
            excelFixKeyDictionary[baseExcel] = new List<int>();
            for (int j = 1; j <= 2; j++)
            {
                var linkExcel = sheet.Cells[1, 1].Offset[16 + (i - 1) * 4, j + 1].Value;
                var baseExcelFixKey = sheet.Cells[1, 1].Offset[17 + (i - 1) * 4, j + 1].Value;
                excelLinkDictionary[baseExcel].Add(linkExcel);
                excelFixKeyDictionary[baseExcel].Add(Convert.ToInt32(baseExcelFixKey));
            }
        }
    }

    public static void test()
    {
        ExcelDic();
        List<string> modeIDRow = new List<string>();
        modeIDRow.Add("10");
        List<string> fileName = new List<string>();
        fileName.Add("索引1.xlsx");
        List<List<string>> modeIDGroup = new List<List<string>>();
        List<string> modeID = new List<string>();
        modeID.Add("####");
        modeIDGroup.Add(modeID);
        var sheet = IndexWk.ActiveSheet;
        string WriteMode = sheet.Range["B11"].value.ToString();
        string testKey = sheet.Range["C11"].value.ToString();
        CreateRellationShip(fileName, modeIDRow, WriteMode, testKey, modeIDGroup);

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

    public static void test2(List<string> oldstr)
    {
        List<string> newstr = new List<string>();
        foreach (var str in oldstr)
        {
            if (str == null) continue;
            if (excelLinkDictionary.ContainsKey(str))
            {
                foreach (var indestr in excelLinkDictionary[str])
                {
                    newstr.Add(indestr);
                }
            }
            Debug.Print(str + "\n" + "\t");

        }
        if (newstr.Count > 0)
        {
            test2(newstr);
        }

    }
    public static void CreateRellationShip(List<string> oldFileName, List<string> oldmodelID, string WriteMode, string testKey, List<List<string>> oldExcelIDGroup)
    {
        List<string> newmodelID = new List<string>();
        List<string> newFileName = new List<string>();
        List<string> newExcelID = new List<string>();
        List<List<string>> newExcelIDGroup = new List<List<string>>();
        int excount = 0;
        foreach (var excelFile in oldFileName)
        {
            var excel = new FileStream(excelPath + @"\" + excelFile, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(excel);
            var sheet = workbook.GetSheetAt(0);
            var rowReSourceRow = FindSourceRow(sheet, 1, oldmodelID[excount]);
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
                var rowTarget = sheet.GetRow(rowReSourceRow + i + 1) ?? sheet.CreateRow(rowReSourceRow + i + 1);
                var rowTargetTemp = sheet.GetRow(rowReSourceRow + i + 1) ?? sheet.CreateRow(rowReSourceRow + i + 1);
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
                            var asdasd = oldExcelIDGroup[i][excount];
                            cellTarget.SetCellValue(asdasd);
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
                    if (excelFileFixKey == 0) continue;
                    var cellTarget = sheet.GetRow(rowReSourceRow).GetCell(excelFileFixKey);
                    var cellTargetValue = ValueTypeToStringInNPOI(cellTarget);
                    //修改字段字典中的字段值，各自方法不一
                    for (int i = 0; i < dataCount; i++)
                    {
                        var rowFix = sheet.GetRow(rowReSourceRow + i + 1) ?? sheet.CreateRow(rowReSourceRow + i + 1);
                        var cellFix = rowFix.GetCell(excelFileFixKey) ?? rowFix.CreateCell(excelFileFixKey);
                        var cellFixValue = ValueTypeToStringInNPOI(cellFix);
                        //每个字段的Value修改方式不一，需要调用方法
                        var cellFixValue2 = cellFixValue + testKey;
                        cellFix.SetCellValue(cellFixValue2);
                        cellFix.CellStyle = cellFix.CellStyle;
                        //有关联表的字段的ID传递出去
                        if (indexExcel != null)
                        {
                            newExcelID.Add(cellFixValue2);
                            //Debug.Print(cellFixValue2+"File=" +excelFile);
                        }
                    }
                    //表格关联字典中寻找下一个递归文件，有关联表的字段ID要生成List递归
                    if (indexExcel != null)
                    {
                        newFileName.Add(indexExcel);
                        newmodelID.Add(cellTargetValue);
                        newExcelIDGroup.Add(newExcelID);
                    }
                    indexExcelCount++;
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
            CreateRellationShip(newFileName, newmodelID, WriteMode, testKey, newExcelIDGroup);
        }
    }
}

