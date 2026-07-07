---
name: numdestools-xlsx-rw
description: Handle xlsx work in NumDesTools. Default to EPPlus for nearly all reads and all writes, formatting, images, merged cells, and formal output files. Use MiniExcel only for explicit high-performance read-only paths such as large scans, global search, cross-file indexing, or diff/conflict analysis where streaming and early break matter.
---

# numdestools-xlsx-rw

Use this skill whenever you touch xlsx logic in NumDesTools.

## Default: EPPlus

Use EPPlus by default for:

- regular reads
- all writes
- styles, borders, row height, column width
- images, merged cells, worksheet structure
- any final xlsx artifact kept by the project

```csharp
using OfficeOpenXml;

ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
using var pkg = new ExcelPackage(new FileInfo(path));
var ws = pkg.Workbook.Worksheets[0];

var text = ws.Cells[row, col].Text?.Trim() ?? "";
ws.Cells[row, col].Value = "文本";
pkg.SaveAs(new FileInfo(outPath));
```

Prefer these project references:

- `NumDesTools.Scanner/ExcelReader.cs`
- `NumDesTools.Scanner/ActivityWriter.cs`
- `NumDesTools.Scanner/LteMapWriter.cs`

## Exception: MiniExcel for explicit performance paths

Use MiniExcel only when the task is dominated by large read/query volume and does not need workbook mutation.

Typical cases:

- full-sheet or multi-file scans
- global search
- index building
- diff / conflict analysis
- stream-until-found flows with early break

```csharp
using MiniExcelLibs;

var rows = MiniExcel.Query(
    filePath,
    sheetName: sheetName,
    configuration: NumDesAddIn.OnOffMiniExcelCatches
).ToList();
```

Rules:

- In the main add-in, pass `NumDesAddIn.OnOffMiniExcelCatches`
- Do not use MiniExcel for writes
- If styles, images, merged cells, or structure matter, use EPPlus
- Do not add new MiniExcel reads without a clear performance reason
- Do not introduce Python-based xlsx generation for formal project output
