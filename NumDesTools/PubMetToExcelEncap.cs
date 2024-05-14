using OfficeOpenXml;
using System.Dynamic;

#pragma warning disable CA1416

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
    public List<dynamic> Read(ExcelWorksheet sheet, int rowFirst, int rowEnd, int colFirst = 1, int indexRow = 4)
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
    public Dictionary<string, List<object>> ReadToDic(ExcelWorksheet sheet, int rowFirst, int colFirst,
        List<int> usedData, int rowEnd = 1)
    {
        Dictionary<string, List<object>> dataDict = new Dictionary<string, List<object>>();
        var colCount = usedData.Count();
        if (rowEnd != 1)
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

    public static Dictionary<string, List<object>> ReadExcelDataToDicByEpplus(string path, string sheetName,
        int startRow, int startCol, List<int> usedData)
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
    public void Write(ExcelWorksheet sheet, ExcelPackage Excel, List<dynamic> data, int rowFirst)
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

    public int FindFromRow(ExcelWorksheet sheet, int col, string searchValue)
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

    public int FindFromCol(ExcelWorksheet sheet, int row, string searchValue)
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

public class ExcelDataByVsto
{
    public dynamic ActiveWorkbook { get; set; }
    public dynamic ActiveSheet { get; set; }
    public string ActiveWorkbookPath { get; set; }

    //创建对象获取基本信息
    public void GetExcelObj()
    {
        dynamic app = NumDesAddIn.App;
        dynamic activeWorkbook = app.ActiveWorkbook;
        dynamic activeSheet = app.ActiveSheet;
        string activeWorkbookPath = activeWorkbook.Path;

        ActiveWorkbook = activeWorkbook;
        ActiveSheet = activeSheet;
        ActiveWorkbookPath = activeWorkbookPath;
    }

    //数据读取
    public (List<object> sheetHeaderCol, List<List<object>> sheetData) Read(Range rangeData, Range rangeHeader,
        int headRow)
    {
        // 读取数据到一个二维数组中
        object[,] rangeValue = rangeData.Value;
        // 读取数据到一个二维数组中
        object[,] headRangeValue = rangeHeader.Value;
        // 定义工作表数据数组和表头数组
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
        // 读取数据
        for (var row = 1; row <= rangeValue.GetLength(0); row++)
        {
            var rowList = new List<object>();
            for (var column = 1; column <= rangeValue.GetLength(1); column++)
            {
                var valueData = rangeValue[row, column];
                rowList.Add(valueData);
            }

            sheetData.Add(rowList);
        }

        //读取表头
        for (var column = 1; column <= rangeValue.GetLength(1); column++)
        {
            var value = headRangeValue[headRow, column];
            sheetHeaderCol.Add(value);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    //通过C-API的方式写入打开当前活动Excel表格各个Sheet的数据
    public void Write(string sheetName, int rowFirst, int rowLast, int colFirst, int colLast,
        object[,] rangeValue)
    {
        var sheet = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);
        var range = new ExcelReference(rowFirst, rowLast, colFirst, colLast, sheet.SheetId);
        ExcelAsyncUtil.QueueAsMacro(() => { range.SetValue(rangeValue); });
    }

    //VSTO内置在Range内查找特定值(第一个)的方法
    public (int row, int column) FindValue(Range searchRange, object valueToFind)
    {
        // 使用 Find 方法在指定范围内查找特定值
        Range foundRange = searchRange.Find(valueToFind);
        // 如果找到了特定值
        if (foundRange != null)
        {
            // 返回找到的单元格的行号和列号
            return (foundRange.Row, foundRange.Column);
        }
        else
        {
            // 如果没有找到特定值，返回 (-1, -1)
            return (-1, -1);
        }
    }
}