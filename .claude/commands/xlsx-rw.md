---
name: xlsx-rw
description: 生成或读取 xlsx 文件。读用 MiniExcel，写用 EPPlus（ExcelPackage）。项目已有这两个库，直接复用，不新增依赖。
---

# xlsx 读写规范（NumDesTools 项目）

## 读取 xlsx — 用 MiniExcel

```csharp
using MiniExcelLibs;

// 读所有 Sheet 名
var sheetNames = MiniExcel.GetSheetNames(filePath);

// 读某 Sheet 所有行（dynamic，首行为列名）
var rows = MiniExcel.Query(filePath, sheetName: sheetName,
    configuration: NumDesAddIn.OnOffMiniExcelCatches).ToList();

foreach (IDictionary<string, object> row in rows)
{
    var val = row["列名"]?.ToString();
}

// 强类型读（需定义 class）
var typed = MiniExcel.Query<MyRow>(filePath, sheetName: sheetName).ToList();
```

**注意**：`NumDesAddIn.OnOffMiniExcelCatches` 是项目全局配置，读时统一传入。

---

## 写入 xlsx — 用 EPPlus（ExcelPackage）

```csharp
using OfficeOpenXml;
using OfficeOpenXml.Style;

// 新建
using var pkg = new ExcelPackage();
// 打开已有文件
using var pkg = new ExcelPackage(new FileInfo(path));

var ws = pkg.Workbook.Worksheets.Add("Sheet1");

// 写单元格
ws.Cells[row, col].Value = "文本";
ws.Cells[row, col].Value = 123;

// 合并
ws.Cells[r1, c1, r2, c2].Merge = true;

// 样式
var cell = ws.Cells[row, col];
cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0x2A, 0x10, 0x50));
cell.Style.Font.Bold = true;
cell.Style.Font.Color.SetColor(System.Drawing.Color.Gold);
cell.Style.Font.Size = 12;
cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
cell.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;
cell.Style.WrapText = true;
var border = cell.Style.Border;
border.Top.Style = border.Bottom.Style =
border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;

// 列宽 / 行高
ws.Column(col).Width = 20;
ws.Row(row).Height = 40;

// 嵌入图片（图片放在指定行列，不遮文字的做法：图片行专用，不写文字）
var pic = ws.Drawings.AddPicture("name", new FileInfo(imgPath));
pic.SetPosition(row - 1, 2, col - 1, 2);   // row/col 从0起，offset px
pic.SetSize(width_px, height_px);

// 保存
pkg.SaveAs(new FileInfo(outPath));
```

### 图片不遮文字的布局原则
- **图片行**：整行只放图片，行高设为图片高度，该行所有文字单元格留空
- **文字行**：不放图片，正常写内容
- 图片用 `SetPosition(rowIndex, rowOffset, colIndex, colOffset)` 定位，rowOffset/colOffset 用小值（2~4px）避免溢出

---

## Python 生成 xlsx（临时分析脚本）— 同样用 EPPlus 风格排版

用 `openpyxl`，图片与文字**严格分行**：

```python
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# 图片专用行：行高 = 图片显示高度，该行文字列全部为空
ws.row_dimensions[img_row].height = IMG_H_PT   # 单位是磅，px * 0.75
img = XLImage(path)
img.width, img.height = W, H
ws.add_image(img, f"A{img_row}")               # 锚定到图片行

# 文字行：紧接在图片行下方
ws.cell(row=txt_row, column=1, value="说明文字")
```

**禁止**：图片锚定行与文字在同一行（openpyxl 图片会浮动覆盖文字）。
