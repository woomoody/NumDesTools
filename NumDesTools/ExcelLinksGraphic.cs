using System.Collections.Generic;
// ReSharper disable All

namespace NumDesTools;

/// <summary>
/// 输出dot文件表现Excel索引关系图
/// </summary>
public class ExcelLinksGraphic
{
    public static void Graph()
    {
        //打开Excel文件
        var workbook = NumDesAddIn.App.ActiveWorkbook;
        var sheet = workbook.ActiveSheet;
        //读取Excel文件统计数据
        var mainExcel = new Dictionary<string, List<string>>();
        var usedRange = sheet.UsedRange;
        for (var row = 1; row <= usedRange.Rows.Count; row++)
        {
            var linkExcel = new List<string>();
            var mainCell = usedRange.Cells[row, 1].Value;
            var mainValue = "";
            if (mainCell != null) mainValue = mainCell.ToString();
            if (mainValue != "")
            {
                for (var col = 2; col <= usedRange.Columns.Count; col++)
                {
                    var linkCell = usedRange.Cells[row, col].Value;
                    var linkValue = "";
                    if (linkCell != null) linkValue = linkCell.ToString() + @".xlsx";
                    if (linkValue != "") linkExcel.Add(linkValue);
                }

                mainExcel[mainValue] = linkExcel;
            }
        }

        //生成关系图
        using (var file = new System.IO.StreamWriter(@"C:\Users\admin\Desktop\output.dot"))
        {
            file.WriteLine("digraph G {");
            foreach (var pair in mainExcel)
            foreach (var field in pair.Value)
                if (mainExcel.ContainsKey(field))
                    file.WriteLine("\"" + pair.Key + "\" -> \"" + field + "\"");
            file.WriteLine("}");
        }
    }
}