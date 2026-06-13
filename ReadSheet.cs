using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

class Program
{
    static void Main()
    {
        EPPlus.ExcelPackage.LicenseContext = EPPlus.LicenseContext.NonCommercial;
        
        var filePath = @"C:\M1Work\Public\Excels\Tables\#【A创新活动】数值.xlsx";
        using var pkg = new EPPlus.ExcelPackage(new System.IO.FileInfo(filePath));
        
        // List all sheet names
        Console.WriteLine("=== Available Sheets ===");
        foreach (var ws in pkg.Workbook.Worksheets)
        {
            Console.WriteLine($"  - {ws.Name}");
        }
        
        // Read the specific sheet
        var targetSheet = "二合棋盘V6-4【高】";
        Console.WriteLine($"\n=== Reading Sheet: {targetSheet} ===\n");
        
        var sheet = pkg.Workbook.Worksheets[targetSheet];
        if (sheet == null)
        {
            Console.WriteLine("Sheet not found!");
            return;
        }
        
        int colCount = sheet.Dimension?.Columns ?? 0;
        int rowCount = sheet.Dimension?.Rows ?? 0;
        Console.WriteLine($"Dimensions: {rowCount} rows x {colCount} columns\n");
        
        // Print all data with row and column numbers
        Console.WriteLine("=== Full Sheet Content ===\n");
        
        // Print column headers
        Console.Write("Row\Col\t");
        for (int c = 1; c <= colCount; c++)
        {
            Console.Write($"C{c:D2}\t");
        }
        Console.WriteLine();
        Console.WriteLine(new string('-', (colCount + 1) * 8));
        
        // Print rows
        for (int r = 1; r <= rowCount; r++)
        {
            Console.Write($"R{r:D3}\t");
            for (int c = 1; c <= colCount; c++)
            {
                var cell = sheet.Cells[r, c];
                var value = cell.Value?.ToString() ?? "";
                Console.Write($"{value}\t");
            }
            Console.WriteLine();
        }
    }
}
