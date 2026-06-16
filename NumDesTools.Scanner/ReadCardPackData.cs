using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

Console.WriteLine("=== 1. 读取 ActivityPhotoCardPack 卡包信息 ===\n");

string file1 = @"C:\M1Work\Public\Excels\Tables\$活动照片收集.xlsx";
var targetIds = new[] { 5600011, 5600012, 5600013, 5600014, 5600015, 5600061 };

try {
    using (var package = new ExcelPackage(new FileInfo(file1)))
    {
        var sheetNames = package.Workbook.Worksheets.Select(s => s.Name).ToList();
        Console.WriteLine($"可用Sheet: {string.Join(", ", sheetNames)}\n");
        
        var sheet = package.Workbook.Worksheets["ActivityPhotoCardPack"];
        if (sheet == null)
        {
            Console.WriteLine("ERROR: 未找到 ActivityPhotoCardPack sheet");
            return;
        }

        // 获取字段名
        var headers = new Dictionary<string, int>();
        for (int col = 1; col <= sheet.Dimension?.Columns; col++)
        {
            var headerVal = sheet.Cells[1, col].Value;
            if (headerVal != null)
                headers[headerVal.ToString()] = col;
        }

        Console.WriteLine($"字段列表({headers.Count}个):");
        foreach (var h in headers.Keys.OrderBy(x => x))
            Console.WriteLine($"  - {h}");
        
        Console.WriteLine();

        // 查找目标ID
        var targetRows = new Dictionary<int, int>();
        for (int row = 2; row <= sheet.Dimension?.Rows; row++)
        {
            var idCol = headers.ContainsKey("id") ? headers["id"] : 1;
            var idCell = sheet.Cells[row, idCol];
            if (idCell.Value != null && int.TryParse(idCell.Value.ToString(), out int id) && targetIds.Contains(id))
                targetRows[id] = row;
        }

        Console.WriteLine($"找到的卡包ID: {string.Join(", ", targetRows.Keys.OrderBy(x => x))}\n");
        Console.WriteLine("=== 详细信息 ===\n");

        // 输出详细信息
        foreach (var id in targetRows.Keys.OrderBy(x => x))
        {
            var row = targetRows[id];
            Console.WriteLine($"--- ID: {id} ---");
            
            foreach (var header in headers.Keys.OrderBy(x => x))
            {
                var col = headers[header];
                var value = sheet.Cells[row, col].Value;
                if (value != null)
                    Console.WriteLine($"  {header}: {value}");
            }
            Console.WriteLine();
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine($"ERROR: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
}

Console.WriteLine("\n=== 2. 读取活动服务端表 ===\n");

string file2 = @"C:\M1Work\Public\Excels\Tables\#活动服务端表ActivityServerData.xlsm";
var targetActivityIds = new[] { 740041, 740050, 940090, 980090, 980160 };

try {
    using (var package = new ExcelPackage(new FileInfo(file2)))
    {
        var sheetNames = package.Workbook.Worksheets.Select(s => s.Name).ToList();
        Console.WriteLine($"可用Sheet({sheetNames.Count}个): {string.Join(", ", sheetNames)}\n");
        
        var mainSheet = package.Workbook.Worksheets.FirstOrDefault();
        if (mainSheet == null)
        {
            Console.WriteLine("ERROR: 无可用Sheet");
            return;
        }

        Console.WriteLine($"主Sheet: {mainSheet.Name}");
        Console.WriteLine($"维度: {mainSheet.Dimension?.Rows} rows x {mainSheet.Dimension?.Columns} cols\n");

        // 获取字段名
        var headers = new Dictionary<string, int>();
        for (int col = 1; col <= mainSheet.Dimension?.Columns; col++)
        {
            var headerVal = mainSheet.Cells[1, col].Value;
            if (headerVal != null)
                headers[headerVal.ToString()] = col;
        }

        // 显示字段名
        var fieldList = headers.Keys.OrderBy(x => x).ToList();
        Console.WriteLine($"字段列表({fieldList.Count}个): {string.Join(", ", fieldList.Take(10))}");
        if (fieldList.Count > 10)
            Console.WriteLine($"... (还有{fieldList.Count - 10}个字段)");
        Console.WriteLine();

        // 查找活动ID列
        int activityIdCol = 0;
        foreach (var (header, col) in headers.OrderBy(x => x.Key))
        {
            if (header.Contains("活动") || header.Contains("activity") || header.Contains("ID") || header == "id")
            {
                activityIdCol = col;
                Console.WriteLine($"检测到活动ID列: 第{col}列 ({header})");
                break;
            }
        }

        if (activityIdCol == 0)
        {
            Console.WriteLine("未找到活动ID列，尝试第1列");
            activityIdCol = 1;
        }
        Console.WriteLine();

        // 统计每个活动的档位数
        var activityRewards = new Dictionary<int, int>();
        var cardPackIds = new HashSet<int> { 5600011, 5600012, 5600013, 5600014, 5600015, 5600061 };

        for (int row = 2; row <= mainSheet.Dimension?.Rows; row++)
        {
            var actValue = mainSheet.Cells[row, activityIdCol].Value;
            if (actValue != null && int.TryParse(actValue.ToString(), out int actId) && targetActivityIds.Contains(actId))
            {
                // 检查该行是否包含卡包ID
                bool hasCardPack = false;
                for (int col = 1; col <= mainSheet.Dimension?.Columns; col++)
                {
                    var cellValue = mainSheet.Cells[row, col].Value;
                    if (cellValue != null && cardPackIds.Any(id => cellValue.ToString().Contains(id.ToString())))
                    {
                        hasCardPack = true;
                        break;
                    }
                }

                if (hasCardPack)
                {
                    if (!activityRewards.ContainsKey(actId))
                        activityRewards[actId] = 0;
                    activityRewards[actId]++;
                }
            }
        }

        Console.WriteLine("=== 活动档位统计 ===\n");
        Console.WriteLine("活动ID | 档位数");
        Console.WriteLine("-------|-------");
        
        foreach (var actId in targetActivityIds.OrderBy(x => x))
        {
            var count = activityRewards.ContainsKey(actId) ? activityRewards[actId] : 0;
            Console.WriteLine($"{actId}  | {count}");
        }
        
        Console.WriteLine();
        Console.WriteLine("统计完成");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"ERROR: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
}
