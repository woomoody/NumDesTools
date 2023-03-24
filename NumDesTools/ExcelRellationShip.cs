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
            excelFixKeyDictionary[baseExcel] =new List<int>();
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
        CreateRellationShip(fileName, modeIDRow);

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

    public static int FindSourceRow(ISheet sheet,int col,string searchValue)
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
            if(str ==null) continue;
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
    public static void CreateRellationShip(List<string> oldFileName, List<string> oldmodelID)
    {
        List<string> newmodelID = new List<string>();
        List<string> newFileName = new List<string>();
        int excount = 0;
        foreach (var excelFile in oldFileName)
        {
            var excel = new FileStream(excelPath + @"\" + excelFile, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(excel);
            var sheet = workbook.GetSheetAt(0);
            var rowReSourceRow = FindSourceRow(sheet, 1, oldmodelID[excount]);
            if(rowReSourceRow==-1) continue;
            var rowSource = sheet.GetRow(rowReSourceRow);
            var colTotal = sheet.GetRow(1).LastCellNum + 1;
            if (sheet.LastRowNum != rowReSourceRow)
            {
                sheet.ShiftRows(rowReSourceRow + 1, sheet.LastRowNum, dataCount, true, false);
            }
            //数据复制
            for (int i = 0; i < dataCount; i++)
            {
                //if (sheet.LastRowNum != rowReSourceRow + i )
                //{
                //    sheet.ShiftRows(rowReSourceRow + i + 1, sheet.LastRowNum, 1, true, false);
                //}
                var rowTarget= sheet.GetRow(rowReSourceRow + i + 1)?? sheet.CreateRow(rowReSourceRow + i + 1);
                for (int j = 0; j < colTotal; j++)
                {
                    var cellSource = rowSource.GetCell(j);
                    var cellSourceValue = "";
                    if (cellSource != null)
                    {
                        cellSourceValue = ValueTypeToStringInNPOI(cellSource);
                        var cellTarget = rowTarget.GetCell(j)??rowTarget.CreateCell(j);
           
                        if (j==1)
                        {
                            cellTarget.SetCellValue(cellSourceValue+"@@@");
                        }
                        else
                        {
                            cellTarget.SetCellValue(cellSourceValue);
                        }
                        cellTarget.CellStyle = cellSource.CellStyle;
                    }
                }
            }
            //表格关联
            if(excelFile==null) continue;
            if (excelLinkDictionary.ContainsKey(excelFile))
            {
                var indexExcelCount = 0;
                foreach (var indexExcel in excelLinkDictionary[excelFile])
                {
                    if (indexExcel == null) continue;
                    //if (excelLinkDictionary.ContainsKey(indexExcel))
                    //{

                    newFileName.Add(indexExcel);
                    var abc = excelFixKeyDictionary[excelFile][indexExcelCount];
                    var cellTarget = sheet.GetRow(rowReSourceRow).GetCell(abc);
                    //if (cellTarget != null)
                    //{
                        var cellTargetValue = ValueTypeToStringInNPOI(cellTarget);
                        //需要加入reg进行ID分析
                        newmodelID.Add(cellTargetValue);
                    //}
                    //数据修改
                    //数据复制
                    for (int i = 0; i < dataCount; i++)
                    {
                        var rowFix = sheet.GetRow(rowReSourceRow + i + 1) ?? sheet.CreateRow(rowReSourceRow + i + 1);
                        var cellFix = rowFix.GetCell(abc)??rowFix.CreateCell(abc);
                        var cellFixValue = ValueTypeToStringInNPOI(cellFix);
                        cellFix.SetCellValue(cellFixValue+"@@@");
                        cellFix.CellStyle = cellFix.CellStyle;
                    }

                    indexExcelCount++;
                }
                //foreach (var newModelIndex in excelFixKeyDictionary[oldFileName[excount]])
                //{
                //    var cellTarget = sheet.GetRow(rowReSourceRow).GetCell(newModelIndex);
                //    if (cellTarget != null)
                //    {
                //        var cellTargetValue = ValueTypeToStringInNPOI(cellTarget);
                //        //需要加入reg进行ID分析
                //        newmodelID.Add(cellTargetValue);
                //    }
                //}

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
            CreateRellationShip(newFileName, newmodelID);
        }
    }
}

