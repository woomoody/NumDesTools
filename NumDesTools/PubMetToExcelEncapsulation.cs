using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;

// ReSharper disable All

namespace NumDesTools;

/// <summary>
/// 公共的Excel功能类-封装为类和属性
/// </summary>
public class ExcelDataByEpplus
{
    public ExcelPackage Excel { get; set; }
    public ExcelWorksheet Sheet { get; set; }
    public List<(string, string, string)> ErrorList { get; private set; }
    public bool GetExcelObj(dynamic excelPath, dynamic excelName)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ExcelWorksheet sheet;
        ExcelPackage excel;
        string errorExcelLog;
        var errorList = new List<(string, string, string)>();
        string path;
        var newPath = Path.GetDirectoryName(Path.GetDirectoryName(excelPath));
        //兼容多表格的工作簿
        string sheetRealName = "Sheet1";
        string excelRealName = excelName;
        if (excelName.Contains("#"))
        {
            var excelRealNameGroup = excelName.Split("#");
            excelRealName = excelRealNameGroup[0];
            sheetRealName = excelRealNameGroup[1];
        }
        switch (excelName)
        {
            case "Localizations.xlsx":
                path = newPath + @"\Excels\Localizations\Localizations.xlsx";
                break;
            case "UIConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIConfigs.xlsx";
                break;
            case "UIItemConfigs.xlsx":
                path = newPath + @"\Excels\UIs\UIItemConfigs.xlsx";
                break;
            default:
                path = excelPath + @"\" + excelRealName;
                break;
        }

        var fileExists = File.Exists(path);
        if (fileExists == false)
        {
            errorExcelLog = excelRealName + "不存在表格文件";
            errorList.Add((excelRealName, errorExcelLog, excelRealName));
            ErrorList = errorList;
            return false;
        }

        excel = new ExcelPackage(new FileInfo(path));
        ExcelWorkbook workBook;
        try
        {
            workBook = excel.Workbook;
        }
        catch (Exception ex)
        {
            errorExcelLog = excelRealName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelRealName, errorExcelLog, excelRealName));
            ErrorList = errorList;
            return false;
        }

        try
        {
            sheet = workBook.Worksheets[sheetRealName];
        }
        catch (Exception ex)
        {
            errorExcelLog = excelRealName + "#不能创建WorkBook对象" + ex.Message;
            errorList.Add((excelRealName, errorExcelLog, excelRealName));
            ErrorList = errorList;
            return false;
        }

        sheet ??= workBook.Worksheets[0];
        ErrorList = errorList;
        Sheet = sheet;
        Excel = excel;
        return true;
    }
    //数据读取
    public List<dynamic> Read(ExcelWorksheet sheet,  int rowFirst, int rowEnd)
    {
        var list = new List<dynamic>();
        int colCount = sheet.Dimension.Columns;
        for (int row = rowFirst; row <= rowEnd; row++)
        {
            var expando = new ExpandoObject() as IDictionary<string, object>;
            for (int col = 1; col <= colCount; col++)
            {
                //索引在第几行
                string columnName = sheet.Cells[4, col].Value?.ToString() ?? string.Empty;
                expando[columnName] = sheet.Cells[row, col].Value;
            }
            list.Add(expando);
        }
        return list;
    }
    //数据写入
    public  void Write(ExcelWorksheet sheet, ExcelPackage Excel ,  List<dynamic> data, int rowFirst)
    {
        // 更新 Excel 数据
        for (int row = 0; row < data.Count; row++)
        {
            var dataRow = (IDictionary<string, object>)data[row];
            int col = 1;
            foreach (var keyValuePair in dataRow)
            {
                sheet.Cells[row + rowFirst, col].Value = keyValuePair.Value;
                col++;
            }
        }
        Excel.Save();
    }
}
