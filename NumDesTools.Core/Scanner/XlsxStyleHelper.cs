using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

internal static class XlsxStyleHelper
{
    internal static void Header(ExcelRange c, string text, string hex = "2F5496")
    {
        c.Value = text;
        c.Style.Fill.PatternType = ExcelFillStyle.Solid;
        c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
        c.Style.Font.Bold = true;
        c.Style.Font.Color.SetColor(Color.White);
        c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        c.Style.WrapText = true;
        Border(c);
    }

    internal static void Cell(ExcelRange c, object? value, string? hex = null, bool wrap = false)
    {
        c.Value = value;
        if (hex is not null)
        {
            c.Style.Fill.PatternType = ExcelFillStyle.Solid;
            c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
        }
        if (wrap)
            c.Style.WrapText = true;
        Border(c);
    }

    internal static void Border(ExcelRange c)
    {
        var b = c.Style.Border;
        b.Top.Style = b.Bottom.Style = b.Left.Style = b.Right.Style = ExcelBorderStyle.Thin;
        var gray = Color.FromArgb(0xBD, 0xBD, 0xBD);
        b.Top.Color.SetColor(gray);
        b.Bottom.Color.SetColor(gray);
        b.Left.Color.SetColor(gray);
        b.Right.Color.SetColor(gray);
    }

    internal static Color HexColor(string hex)
    {
        hex = hex.TrimStart('#');
        return Color.FromArgb(
            Convert.ToInt32(hex[..2], 16),
            Convert.ToInt32(hex[2..4], 16),
            Convert.ToInt32(hex[4..6], 16)
        );
    }

    internal static void MechNote(
        ExcelWorksheet ws,
        int row,
        int startCol,
        string text,
        int mergeWidth
    )
    {
        var cell = ws.Cells[row, startCol, row, startCol + mergeWidth - 1];
        cell.Merge = true;
        cell.Value = text;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(HexColor("F0F4FA"));
        cell.Style.Font.Size = 9;
        cell.Style.WrapText = true;
        Border(cell);
    }
}
