using System.Collections.Generic;

namespace NumDesTools;
/// <summary>
/// 输出dot文件表现Excel索引关系图
/// </summary>
public class ExcelLinksGraphic
{
    public static void Graph()
    {
        //打开Excel文件
        var workbook = CreatRibbon.App.ActiveWorkbook;
        var sheet = workbook.ActiveSheet;
        //读取Excel文件统计数据
        Dictionary<string, List<string>> mainExcel = new Dictionary<string, List<string>>();
        var usedRange = sheet.UsedRange;
        for (int row = 1; row <= usedRange.Rows.Count; row++)
        {
            List<string> linkExcel = new List<string>();
            var mainCell = usedRange.Cells[row, 1].Value;
            string mainValue = "";
            if (mainCell != null)
            {
                mainValue = mainCell.ToString();
            }
            if (mainValue!= "")
            {
                for (int col = 2; col <= usedRange.Columns.Count; col++)
                {
                    var linkCell = usedRange.Cells[row, col].Value;
                    string linkValue = "";
                    if (linkCell != null)
                    {
                        linkValue = linkCell.ToString() + @".xlsx";
                    }
                    if (linkValue != "")
                    {
                        linkExcel.Add(linkValue);
                    }
                }
                mainExcel[mainValue] = linkExcel;
            }
        }
        //生成关系图
        using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\admin\Desktop\output.dot"))
        {
            file.WriteLine("digraph G {");
            foreach (KeyValuePair<string, List<string>> pair in mainExcel)
            {
                foreach (string field in pair.Value)
                {
                    if (mainExcel.ContainsKey(field))
                    {
                        file.WriteLine("\"" + pair.Key + "\" -> \"" + field + "\"");
                    }
                }
            }
            file.WriteLine("}");
        }
    }

}