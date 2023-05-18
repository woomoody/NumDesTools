using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace NumDesTools;

public class FlowChartEditor
{
    private const string ExcelFilePath = "流程图.xlsx";
    private const string TextFilePath = "流程图.txt";

    public void CreateFlowChart()
    {
        Excel.Application excelApp = new Excel.Application();
        excelApp.Visible = true;

        Excel.Workbook workbook = excelApp.Workbooks.Add();
        Excel.Worksheet worksheet = workbook.ActiveSheet as Excel.Worksheet;

        worksheet.Cells[1, 1].Value = "开始";
        worksheet.Cells[2, 1].Value = "步骤1";
        worksheet.Cells[2, 2].Value = "步骤2";
        worksheet.Cells[3, 1].Value = "结束";

        workbook.SaveAs(ExcelFilePath);

        workbook.Close();
        excelApp.Quit();

        Console.WriteLine("流程图已创建并保存为Excel文件。");
    }

    public void EditFlowChart()
    {
        if (!File.Exists(ExcelFilePath))
        {
            Console.WriteLine("找不到流程图的Excel文件。请先创建流程图。");
            return;
        }

        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = excelApp.Workbooks.Open(ExcelFilePath);
        Excel.Worksheet worksheet = workbook.ActiveSheet as Excel.Worksheet;

        Console.WriteLine("请输入流程图的步骤，每行一个步骤，输入空行表示结束编辑：");
        List<string> flowChart = new List<string>();
        string step = Console.ReadLine();
        while (!string.IsNullOrEmpty(step))
        {
            flowChart.Add(step);
            step = Console.ReadLine();
        }

        for (int i = 1; i <= flowChart.Count; i++)
        {
            worksheet.Cells[i, 1].Value = flowChart[i - 1];
        }

        workbook.Save();

        workbook.Close();
        excelApp.Quit();

        SaveFlowChartToTextFile(flowChart);

        Console.WriteLine("流程图已编辑并保存为Excel文件和文本文件。");
    }

    public void RestoreFlowChart()
    {
        List<string> flowChart = LoadFlowChartFromTextFile();
        if (flowChart.Count == 0)
        {
            Console.WriteLine("找不到流程图的文本文件。请先编辑流程图。");
            return;
        }

        Console.WriteLine("还原的流程图内容如下：");
        foreach (string step in flowChart)
        {
            Console.WriteLine(step);
        }

        Console.WriteLine("流程图还原完成。");
    }

    private void SaveFlowChartToTextFile(List<string> flowChart)
    {
        File.WriteAllLines(TextFilePath, flowChart);
    }

    private List<string> LoadFlowChartFromTextFile()
    {
        if (!File.Exists(TextFilePath))
        {
            return new List<string>();
        }

        return new List<string>(File.ReadAllLines(TextFilePath));
    }
}
public class Program
{
    public static void NodeMain()
    {
        FlowChartEditor editor = new FlowChartEditor();
        editor.CreateFlowChart();
        //Console.WriteLine("1. 创建流程图");
        //Console.WriteLine("2. 编辑流程图");
        //Console.WriteLine("3. 还原流程图");
        //Console.WriteLine("0. 退出");

        //while (true)
        //{
        //    Console.WriteLine("请选择操作：");
        //    string option = Console.ReadLine();
        //    switch (option)
        //    {
        //        case "1":
        //            editor.CreateFlowChart();
        //            break;
        //        case "2":
        //            editor.EditFlowChart();
        //            break;
        //        case "3":
        //            editor.RestoreFlowChart();
        //            break;
        //        case "0":
        //            return;
        //        default:
        //            Console.WriteLine("无效的选项，请重新选择。");
        //            break;
        //    }
        //}
    }
}