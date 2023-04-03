//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using IExcel = Microsoft.Office.Interop.Excel;

using ExcelDna.Integration;
using NPOI.XSSF.UserModel;
using System.IO;

namespace NumDesTools;

public class ExcelUdf
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic IndexWk = App.ActiveWorkbook;
    private static readonly dynamic excelPath = IndexWk.Path;
    [ExcelFunction(Category = "test", IsVolatile = true, IsMacroType = true, Description = "测试自定义函数")]
    public static double Sum2Num([ExcelArgument(Description = "选个格子")] double a,
        [ExcelArgument(Description = "选个格子")] double b)
    {
        return a + b;
    }
    [ExcelFunction(Category = "test2", IsVolatile = true, IsMacroType = true, Description = "寻找指定表格字段所在列")]
    public static int FindKeyCol([ExcelArgument(Description = "工作簿")] string targetWorkbook, [ExcelArgument(Description = "目标行")] int row, [ExcelArgument(Description = "匹配值")] string searchValue, [ExcelArgument(Description = "工作表")] string targetSheet ="Sheet1")
    {
        var path = excelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read,FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet);
        if (sheet == null)
        {
            sheet= workbook.GetSheetAt(0);
        }
        var rowSource = sheet.GetRow(row);
        for (int j = rowSource.FirstCellNum; j <= rowSource.LastCellNum; j++)
        {
            var cell = rowSource.GetCell(j);
            if (cell != null)
            {
                var cellValue = cell.ToString();
                if (cellValue == searchValue)
                {
                    workbook.Close();
                    fs.Close();
                    return j;
                }
            }
        }
        workbook.Close();
        fs.Close();
        return -1;
    }
    [ExcelFunction(Category = "test2", IsVolatile = true, IsMacroType = true, Description = "寻找指定表格字段所在列")]
    public static int FindKeyRow([ExcelArgument(Description = "工作簿")] string targetWorkbook, [ExcelArgument(Description = "目标列")] int col, [ExcelArgument(Description = "匹配值")] string searchValue, [ExcelArgument(Description = "工作表")] string targetSheet="Sheet1")
    {
        var path = excelPath + @"\" + targetWorkbook;
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read,FileShare.ReadWrite);
        var workbook = new XSSFWorkbook(fs);
        var sheet = workbook.GetSheet(targetSheet);
        if (sheet == null)
        {
            sheet = workbook.GetSheetAt(0);
        }
        for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
        {
            var rowSource = sheet.GetRow(i);
            if (rowSource != null)
            {
                var cell = rowSource.GetCell(col);
                var cellValue = cell.ToString();
                if (cellValue == searchValue)
                {
                    workbook.Close();
                    fs.Close();
                    return i;
                }
            }
        }
        workbook.Close();
        fs.Close();
        return -1;
    }
}