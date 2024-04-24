using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;

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
        if (excelName.Contains("##"))
        {
            var excelRealNameGroup = excelName.Split("##");
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
    public List<dynamic> Read(ExcelWorksheet sheet,  int rowFirst, int rowEnd, int colFirst = 1, int indexRow = 4)
    {
        var list = new List<dynamic>();
        int colCount = sheet.Dimension.Columns;
        for (int row = rowFirst; row <= rowEnd; row++)
        {
            var expando = new ExpandoObject() as IDictionary<string, object>;
            for (int col = colFirst; col <= colCount; col++)
            {
                //索引在第几行
                string columnName = sheet.Cells[indexRow, col].Value?.ToString() ?? string.Empty;
                expando[columnName] = sheet.Cells[row, col].Value;
            }
            list.Add(expando);
        }
        return list;
    }
    //数据读取
    public Dictionary<string, List<object>> ReadToDic(ExcelWorksheet sheet, int rowFirst, int colFirst, List<int> usedData, int rowEnd = 1 )
    {
        Dictionary<string, List<object>> dataDict = new Dictionary<string, List<object>>();
        var colCount = usedData.Count();
        if(rowEnd != 1)
        {
            rowEnd = sheet.Dimension.End.Row;
        }
        string lastMainTable = null;
        for (int i = rowFirst; i <= rowEnd; i++)
        {
            string mainTable = sheet.Cells[i, colFirst].Text;
            if (!string.IsNullOrEmpty(mainTable))
            {
                // 如果"主表"列有值，记住这个值
                lastMainTable = mainTable;
                // 同时为这个主表在字典中创建一个新的内部字典
                if (!dataDict.ContainsKey(lastMainTable))
                {
                    dataDict[lastMainTable] = new List<object>();
                }
            }
            string data;
            List<string> usedDataList = new List<string>();
            for (int j = 0; j < colCount; j++)
            {
                data = sheet.Cells[i, usedData[j]].Text;
                usedDataList.Add(data);
            }
            dataDict[lastMainTable].Add(usedDataList);
        }
        return dataDict;
    }
    public static Dictionary<string, List<object>> ReadExcelDataToDicByEpplus(string path, string sheetName, int startRow, int startCol, List<int> usedData)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Dictionary<string, List<object>> dataDict = new Dictionary<string, List<object>>();
        using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
        {
            //ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // 获取第一个工作表
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName]; // 获取指定工作表
            var colCount = usedData.Count();
            string lastMainTable = null;
            for (int i = startRow; i <= worksheet.Dimension.End.Row; i++)
            {
                string mainTable = worksheet.Cells[i, startCol].Text;
                if (!string.IsNullOrEmpty(mainTable))
                {
                    // 如果"主表"列有值，记住这个值
                    lastMainTable = mainTable;
                    // 同时为这个主表在字典中创建一个新的内部字典
                    if (!dataDict.ContainsKey(lastMainTable))
                    {
                        dataDict[lastMainTable] = new List<object>();
                    }
                }
                string data;
                List<string> usedDataList = new List<string>();
                for (int j = 0; j < colCount; j++)
                {
                    data = worksheet.Cells[i, usedData[j]].Text;
                    usedDataList.Add(data);
                }
                dataDict[lastMainTable].Add(usedDataList);
            }
        }
        return dataDict;
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
    public  int FindFromRow(ExcelWorksheet sheet, int col, string searchValue)
    {
        for (var row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            // 获取当前行的单元格数据
            var cellValue = sheet.Cells[row, col].Value;

            // 如果找到了匹配的值
            if (cellValue != null && cellValue.ToString() == searchValue)
            {
                // 返回该单元格的行地址
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Row;
                return rowAddress;
            }
        }
        return -1;
    }
    public  int FindFromCol(ExcelWorksheet sheet, int row, string searchValue)
    {
        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
        {
            // 获取当前行的单元格数据
            var cellValue = sheet.Cells[row, col].Value;

            // 如果找到了匹配的值
            if (cellValue != null && cellValue.ToString() == searchValue)
            {
                // 返回该单元格的行地址
                var cellAddress = new ExcelCellAddress(row, col);
                var rowAddress = cellAddress.Column;
                return rowAddress;
            }
        }

        return -1;
    }
}
