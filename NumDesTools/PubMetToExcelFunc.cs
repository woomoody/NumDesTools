using ExcelDna.Integration;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;

namespace NumDesTools;
/// <summary>
/// 公共的Excel功能类调用的具体业务逻辑
/// </summary>
public class PubMetToExcelFunc
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    private static readonly dynamic Wk = App.ActiveWorkbook;
    private static readonly dynamic Path = Wk.Path;
    //Excel数据查询并合并表格数据
    public static void ExcelDataSearchAndMerge(string searchValue)
    {
        //获取所有的表格路径
        var rootPath = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(Path));
        var fileList = new List<string>() { rootPath+ @"\Excels\Tables\", rootPath + @"\Excels\Localizations\", rootPath + @"\Excels\UIs\" };
        var files = PubMetToExcel.PathExcelFileCollect(fileList, "*.xlsx", "#");
        //查找指定关键词，记录行号和表格索引号
        var findValueList = new List<(string, string, int, int,string,string)>();
        Parallel.ForEach(files, file =>
        {
            var dataTable = PubMetToExcel.ExcelDataToDataTableOleDb(file);
            var findValue = PubMetToExcel.FindDataInDataTable(file , dataTable, searchValue);
            if (findValue.Count > 0)
            {
                //findValueList.Add(findValue);
                findValueList = findValueList.Concat(findValue).ToList();
            }
        }
            );
        //人工查询所需要的数据，可以打开表格，可以删除和手动增加数据，专用表格进行操作
        dynamic tempWorkbook;
        try
        {
            tempWorkbook = App.Workbooks.Open(rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
        }
        catch
        {
            tempWorkbook = App.Workbooks.Add();
            tempWorkbook.SaveAs(rootPath + @"\Excels\Tables\#合并表格数据缓存.xlsx");
        }
        dynamic tempSheet = tempWorkbook.Sheets["Sheet1"];
        string[,] tempDataArray = new string[findValueList.Count, 5];
        for (int i = 0; i < findValueList.Count; i++)
        {
            tempDataArray[i, 0] = findValueList[i].Item1;
            tempDataArray[i, 1] = findValueList[i].Item2;
            tempDataArray[i, 2] = PubMetToExcel.ConvertToExcelColumn(findValueList[i].Item4)+findValueList[i].Item3;
            tempDataArray[i, 3] = findValueList[i].Item5;
            tempDataArray[i, 4] = findValueList[i].Item6;
            
        }
        var tempDataRange = tempSheet.Range[tempSheet.Cells[2, 2], tempSheet.Cells[2 + tempDataArray.GetLength(0) - 1, 2 + tempDataArray.GetLength(1) - 1]];
        tempDataRange.Value = tempDataArray;
        tempWorkbook.Save();
        //合并数据
    }
    //Excel右键识别文件路径并打开
    public static void RightOpenExcelByActiveCell(CommandBarButton ctrl, ref bool cancelDefault)
    {
        var sheet = App.ActiveSheet;
        var selectCell = App.ActiveCell;
        string selectCellValue = "";
        if (selectCell.Value != null)
        {
            selectCellValue = selectCell.Value.ToString();
        }
        //正则出是Excel路径的单元格
        var isMatch = Regex.IsMatch(selectCellValue, @"^[A-Za-z]:(\\[\w-]+)+(\.xlsx)$");
        if (isMatch)
        {
            var selectRow = selectCell.Row;
            var selectCol = selectCell.Column;
            var sheetName = sheet.Cells[selectRow, selectCol+1].Value;
            var cellAdress = sheet.Cells[selectRow, selectCol + 2].Value;
            PubMetToExcel.OpenExcelAndSelectCell(selectCellValue,sheetName,cellAdress);
        }
    }


    ~PubMetToExcelFunc()
    {
        App.Dispose();
    }

}
