using OfficeOpenXml;
using System;
using System.IO;

namespace NumDesTools.Scanner;

public static class ReadExcelUtil
{
    public static void ReadExcelRanges()
    {

        string filePath = @"C:\M1Work\Public\Excels\Tables\#【A创新活动】数值.xlsx";
        string sheetName = "二合棋盘V6-4【高】";

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[sheetName];
            
            if (worksheet == null)
            {
                Console.WriteLine($"Sheet not found: {sheetName}");
                Console.WriteLine("Available sheets:");
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"  - {ws.Name}");
                }
                return;
            }
            
            Console.WriteLine("=== 第57到110行 ===");
            for (int row = 57; row <= 110; row++)
            {
                for (int col = 1; col <= 26; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    if (cell.Value != null)
                    {
                        Console.WriteLine($"R{row}C{col}={cell.Value}");
                    }
                }
            }
            
            Console.WriteLine("");
            Console.WriteLine("=== 第141到200行 ===");
            for (int row = 141; row <= 200; row++)
            {
                for (int col = 1; col <= 26; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    if (cell.Value != null)
                    {
                        Console.WriteLine($"R{row}C{col}={cell.Value}");
                    }
                }
            }
        }
    }
}
