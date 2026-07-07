---
description: 在 NumDesTools 项目中处理 xlsx 文件时使用此规范。默认使用 EPPlus 读写和结构化读取；仅当任务属于大量查询、全表扫描、跨文件索引、diff/冲突分析等明确的高性能只读路径时，才使用 MiniExcel。禁止为本项目新增 Python 生成 xlsx 路径。
---

# xlsx 读写规范（NumDesTools 项目）

## 默认方案：EPPlus

大多数 xlsx 任务都用 EPPlus，包括：

- 常规读取
- 写入单元格
- 样式、边框、行高列宽
- 图片、合并单元格、工作表结构操作
- 任何最终要落盘的正式 xlsx 产物

```csharp
using OfficeOpenXml;
using OfficeOpenXml.Style;

ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
using var pkg = new ExcelPackage(new FileInfo(path));
var ws = pkg.Workbook.Worksheets[0];

var text = ws.Cells[row, col].Text?.Trim() ?? "";
ws.Cells[row, col].Value = "文本";
ws.Column(col).Width = 20;
ws.Row(row).Height = 40;

pkg.SaveAs(new FileInfo(outPath));
```

优先参考：

- `NumDesTools.Scanner/ExcelReader.cs`
- `NumDesTools.Scanner/ActivityWriter.cs`
- `NumDesTools.Scanner/LteMapWriter.cs`

## 仅限高性能只读路径：MiniExcel

只有在下面这类任务里才用 MiniExcel：

- 大量查询 / 全表扫描
- 跨多个 xlsx 建索引
- 全局搜索、关键词定位
- diff / 冲突分析
- 找到目标后可以提前 break 的流式读取

```csharp
using MiniExcelLibs;

var sheetNames = MiniExcel.GetSheetNames(filePath);
var rows = MiniExcel.Query(
    filePath,
    sheetName: sheetName,
    configuration: NumDesAddIn.OnOffMiniExcelCatches
).ToList();
```

规则：

- 主插件里统一传 `NumDesAddIn.OnOffMiniExcelCatches`
- MiniExcel 只用于读取，不用于写入
- 如果任务需要样式、图片、合并单元格或结构信息，就回到 EPPlus
- 没有明确性能理由时，不要新增 MiniExcel 读取路径

## 禁止项

- 不要用 Python 为本项目生成正式 xlsx
- 不要把“读很多”之外的理由包装成性能需求去引入 MiniExcel
- 不要为写入场景引入 MiniExcel
