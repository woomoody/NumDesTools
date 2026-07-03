using OfficeOpenXml;

namespace NumDesTools.Tests;

public class CellValueNormalizerTests
{
    static CellValueNormalizerTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    // ── Normalize：能安全转数字的才转，转不安全（.Text 会变）的必须留原文本 ──

    [Fact]
    public void Normalize_PlainInteger_ReturnsLong()
    {
        Assert.Equal(350382L, CellValueNormalizer.Normalize("350382"));
    }

    [Fact]
    public void Normalize_NegativeInteger_ReturnsLong()
    {
        Assert.Equal(-123L, CellValueNormalizer.Normalize("-123"));
    }

    [Fact]
    public void Normalize_Zero_ReturnsLong()
    {
        Assert.Equal(0L, CellValueNormalizer.Normalize("0"));
    }

    [Fact]
    public void Normalize_LeadingZeroInteger_StaysNull()
    {
        // "007" 转数字会变成 7，前导零信息丢失，必须保留原文本
        Assert.Null(CellValueNormalizer.Normalize("007"));
    }

    [Fact]
    public void Normalize_LeadingZeroDecimal_ReturnsDouble()
    {
        // "0.5" 的前导零是小数写法本身，不是要保留的 padding，可以安全转
        Assert.Equal(0.5, CellValueNormalizer.Normalize("0.5"));
    }

    [Fact]
    public void Normalize_PlainDecimal_ReturnsDouble()
    {
        Assert.Equal(3.14, CellValueNormalizer.Normalize("3.14"));
    }

    [Fact]
    public void Normalize_TrailingZeroDecimal_StaysNull()
    {
        // "10.0" 若转成 double 用 0.## 格式回显会变成 "10"（丢了 .0），必须保留原文本
        Assert.Null(CellValueNormalizer.Normalize("10.0"));
    }

    [Fact]
    public void Normalize_ScientificNotationString_StaysNull()
    {
        // 字符串本身就是 "1e5"，转数字后语义/显示都会变
        Assert.Null(CellValueNormalizer.Normalize("1e5"));
    }

    [Fact]
    public void Normalize_LongIntegerWithinLongRange_ReturnsLong()
    {
        // 16 位数字在 long(19位精确)范围内，往返无损，应该转
        Assert.Equal(1234567890123456L, CellValueNormalizer.Normalize("1234567890123456"));
    }

    [Fact]
    public void Normalize_HugeNumberBeyondSafeRoundTrip_StaysNull()
    {
        // 超出 long 范围会掉到 double 分支，精度不够往返对不上，必须保留原文本
        Assert.Null(CellValueNormalizer.Normalize("99999999999999999999"));
    }

    [Fact]
    public void Normalize_PlainText_StaysNull()
    {
        Assert.Null(CellValueNormalizer.Normalize("道具"));
    }

    [Fact]
    public void Normalize_ThousandsSeparator_StaysNull()
    {
        Assert.Null(CellValueNormalizer.Normalize("1,000"));
    }

    [Fact]
    public void Normalize_EmptyOrWhitespace_StaysNull()
    {
        Assert.Null(CellValueNormalizer.Normalize(""));
        Assert.Null(CellValueNormalizer.Normalize("   "));
        Assert.Null(CellValueNormalizer.Normalize(null));
    }

    [Fact]
    public void Normalize_TrimsSurroundingWhitespace()
    {
        Assert.Equal(42L, CellValueNormalizer.Normalize("  42  "));
    }

    // ── ApplyTo：写入 cell 后，Text 显示必须跟原字符串一模一样 ────────────────

    [Fact]
    public void ApplyTo_ConvertibleInteger_WritesNumberAndLocksFormat()
    {
        using var pkg = new ExcelPackage();
        var sheet = pkg.Workbook.Worksheets.Add("Sheet1");
        var cell = sheet.Cells[1, 1];

        CellValueNormalizer.ApplyTo(cell, "350382");

        Assert.IsNotType<string>(cell.Value); // 确认落地成了数值类型，不再是字符串
        Assert.Equal("350382", cell.Text);
    }

    [Fact]
    public void ApplyTo_LargeIntegerBeyondElevenDigits_DoesNotBecomeScientificNotation()
    {
        // 复现实测踩过的坑：11 位以上整数在 General 格式下会被 Excel 显示成科学计数法
        using var pkg = new ExcelPackage();
        var sheet = pkg.Workbook.Worksheets.Add("Sheet1");
        var cell = sheet.Cells[1, 1];

        CellValueNormalizer.ApplyTo(cell, "76330182300");

        Assert.Equal("76330182300", cell.Text);
    }

    [Fact]
    public void ApplyTo_NonConvertibleText_KeepsOriginalString()
    {
        using var pkg = new ExcelPackage();
        var sheet = pkg.Workbook.Worksheets.Add("Sheet1");
        var cell = sheet.Cells[1, 1];

        CellValueNormalizer.ApplyTo(cell, "道具");

        Assert.Equal("道具", cell.Value);
        Assert.Equal("道具", cell.Text);
    }

    [Fact]
    public void ApplyTo_LeadingZeroCode_KeepsOriginalString()
    {
        using var pkg = new ExcelPackage();
        var sheet = pkg.Workbook.Worksheets.Add("Sheet1");
        var cell = sheet.Cells[1, 1];

        CellValueNormalizer.ApplyTo(cell, "0512");

        Assert.Equal("0512", cell.Value);
        Assert.Equal("0512", cell.Text);
    }
}
