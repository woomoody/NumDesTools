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
         List<int> modeIDRow = new List<int>();
         modeIDRow.Add(5);
         List<string> fileName = new List<string>();
          fileName.Add("索引1.xlsx"); 
         CreateRellationShip(fileName, modeIDRow);
    }
    public static void CreateRellationShip(List<string> fileName, List<int> modeIDRow)
    {
        List<int> modelDrow2 = new List<int>();
        List<string> fileName2 = new List<string>();
        int excount = 0;
        foreach (var item in modeIDRow)
        {
            var excel = new FileStream(excelPath + @"\" + fileName[excount], FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(excel);
            var sheet = workbook.GetSheetAt(0);
            var rowSource = sheet.GetRow(item);
            var colTotal = sheet.GetRow(1).LastCellNum + 1;
            
            //数据复制
            for (int i = 0; i < dataCount; i++)
            {
                var rowTarget = sheet.GetRow(item + i);
                for (int j = 0; j < colTotal; j++)
                {
                    var cellSource = rowSource.GetCell(j);
                    var cellTarget = rowTarget.CreateCell(j);
                    cellTarget.SetCellValue("cellSource.StringCellValue");
                }
            }
            //数据修改
            if (!excelLinkDictionary.ContainsKey(fileName[excount]))
            {
                var excel2 = new FileStream(excelPath + @"\" + fileName[excount], FileMode.Create, FileAccess.Write);
                workbook.Write(excel2);
                workbook.Close();
                excel2.Close();
                excel.Close();
                continue;
            }
            else
            {
                foreach (var indexExcel in excelLinkDictionary[fileName[excount]])
                {
                    if (indexExcel == null) continue;
                    //if (excelLinkDictionary.ContainsKey(indexExcel))
                    //{
                    fileName2.Add(indexExcel);
                    modelDrow2.Add(4);
                    //}
                }
                var excel2 = new FileStream(excelPath + @"\" + fileName[excount], FileMode.Create, FileAccess.Write);
                workbook.Write(excel2);
                workbook.Close();
                excel2.Close();
                excel.Close();
                excount++;
                if (modelDrow2.Count > 0)
                {
                    CreateRellationShip(fileName2, modelDrow2);
                }
            }
        }
    }
}

