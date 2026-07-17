using OfficeOpenXml;

namespace NumDesTools;

// xlsx 里大量数字被误存成字符串(EPPlus 对 string 类型的 Value 一律走 sharedStrings，
// 数字大多唯一，共享去重零收益，纯体积浪费)。这里把"能安全转数字"的判定和"转后怎么显示"
// 收拢成一处：判定用往返校验(格式化回文本必须跟原字符串完全一致)，天然规避前导零/超长
// ID精度损失/科学计数法这几类会改变显示的转换；写入时顺带锁 NumberFormat，防止大数字
// 在默认 General 格式下被 Excel 显示成科学计数法(实测踩过这个坑)。
internal static class CellValueNormalizer
{
    private const string DoubleFormat = "0.##############";

    // 纯数字字符串 -> 原生数值(long/double)；判定不安全(.Text 会变)的返回 null，原样保留字符串。
    internal static object? Normalize(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;
        var s = raw.Trim();

        if (
            long.TryParse(s, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var l)
            && l.ToString(CultureInfo.InvariantCulture) == s
        )
            return l;

        if (
            double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)
            && d.ToString(DoubleFormat, CultureInfo.InvariantCulture) == s
        )
            return d;

        return null;
    }

    // 把 raw 归一化后写入 cell：能转数字就转数字+锁定对应显示格式，不能转就原样写字符串。
    internal static void ApplyTo(ExcelRange cell, string? raw)
    {
        switch (Normalize(raw))
        {
            case long l:
                cell.Value = l;
                cell.Style.Numberformat.Format = "0";
                break;
            case double d:
                cell.Value = d;
                cell.Style.Numberformat.Format = DoubleFormat;
                break;
            default:
                cell.Value = raw;
                break;
        }
    }
}
