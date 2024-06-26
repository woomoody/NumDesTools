using System.Dynamic;
using OfficeOpenXml;

#pragma warning disable CA1416

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
        string sheetRealName = "Sheet1";
        string excelRealName = excelName;
        if (excelName.Contains("##"))
        {
            var excelRealNameGroup = excelName.Split("##");
            if (excelRealNameGroup.Length == 3)
            {
                excelRealName = excelRealNameGroup[1];
                sheetRealName = excelRealNameGroup[2];
            }
            else
            {
                excelRealName = excelRealNameGroup[0];
                sheetRealName = excelRealNameGroup[1];
            }
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

    public List<dynamic> Read(
        ExcelWorksheet sheet,
        int rowFirst,
        int rowEnd,
        int colFirst = 1,
        int indexRow = 4
    )
    {
        var list = new List<dynamic>();
        int colCount = sheet.Dimension.Columns;
        for (int row = rowFirst; row <= rowEnd; row++)
        {
            var expando = new ExpandoObject() as IDictionary<string, object>;
            for (int col = colFirst; col <= colCount; col++)
            {
                string columnName = sheet.Cells[indexRow, col].Value?.ToString() ?? string.Empty;
                expando[columnName] = sheet.Cells[row, col].Value;
            }

            list.Add(expando);
        }

        return list;
    }

    public Dictionary<string, List<object>> ReadToDic(
        ExcelWorksheet sheet,
        int rowFirst,
        int colFirst,
        List<int> usedData,
        int rowEnd = 1
    )
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
                lastMainTable = mainTable;
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

    public static Dictionary<string, List<object>> ReadExcelDataToDicByEpplus(
        string path,
        string sheetName,
        int startRow,
        int startCol,
        List<int> usedData
    )
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Dictionary<string, List<object>> dataDict = new Dictionary<string, List<object>>();
        using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
            var colCount = usedData.Count();
            string lastMainTable = null;
            for (int i = startRow; i <= worksheet.Dimension.End.Row; i++)
            {
                string mainTable = worksheet.Cells[i, startCol].Text;
                if (!string.IsNullOrEmpty(mainTable))
                {
                    lastMainTable = mainTable;
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

    public void Write(ExcelWorksheet sheet, ExcelPackage excel, List<dynamic> data, int rowFirst)
    {
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

        excel.Save();
    }

    public int FindFromRow(ExcelWorksheet sheet, int col, string searchValue)
    {
        for (var row = 2; row <= sheet.Dimension.End.Row; row++)
        {
            var cellValue = sheet.Cells[row, col].Value;

            if (cellValue != null && cellValue.ToString() == searchValue)
            {
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
            var cellValue = sheet.Cells[row, col].Value;

            if (cellValue != null && cellValue.ToString() == searchValue)
            {
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

    public (List<object> sheetHeaderCol, List<List<object>> sheetData) Read(
        Range rangeData,
        Range rangeHeader,
        int headRow
    )
    {
        object[,] rangeValue = rangeData.Value2;
        object[,] headRangeValue = rangeHeader.Value2;
        var sheetData = new List<List<object>>();
        var sheetHeaderCol = new List<object>();
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

        for (var column = 1; column <= rangeValue.GetLength(1); column++)
        {
            var value = headRangeValue[headRow, column];
            sheetHeaderCol.Add(value);
        }

        var excelData = (sheetHeaderCol, sheetData);
        return excelData;
    }

    public void Write(
        string sheetName,
        int rowFirst,
        int rowLast,
        int colFirst,
        int colLast,
        object[,] rangeValue
    )
    {
        var sheet = (ExcelReference)XlCall.Excel(XlCall.xlSheetId, sheetName);
        var range = new ExcelReference(rowFirst, rowLast, colFirst, colLast, sheet.SheetId);
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            range.SetValue(rangeValue);
        });
    }

    public (int row, int column) FindValue(Range searchRange, object valueToFind)
    {
        Range foundRange = searchRange.Find(valueToFind);
        if (foundRange != null)
        {
            return (foundRange.Row, foundRange.Column);
        }
        else
        {
            return (-1, -1);
        }
    }
}
