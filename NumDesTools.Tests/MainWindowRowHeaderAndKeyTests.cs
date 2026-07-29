using System.Reflection;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// #4 行号映射 + #7 焦点判定的纯逻辑单测（反射调 MainWindow internal static 方法，沿用
/// <see cref="MainWindowBehaviorBaselineTests"/> 加载 dll 反射的既有约定，避免构造 WPF Window）。
/// </summary>
public sealed class MainWindowRowHeaderAndKeyTests
{
    private static readonly Type MainWindowType = LoadType("NumDesTools.XlsxEditor.MainWindow");

    // ── #4：RowHeaderNumber = RowView.RowIndex + 1（绝对 Excel 行号），冻结/主区都对齐 ──

    [Fact]
    public void RowHeaderNumber_UsesRowViewAbsoluteRowIndexPlusOne()
    {
        var store = ColumnStore.Create(["A"], 10);
        for (var r = 0; r < 10; r++)
            store.AppendRow();

        // 冻结区 RowRangeView(store,0,3)：第 0 个 RowView 是 store 行 0 → 行号 1
        var frozenView = new RowRangeView(store, 0, 3);
        Assert.Equal(1, InvokeRowHeaderNumber(frozenView[0], 0)); // 冻结第一行 → 1
        Assert.Equal(3, InvokeRowHeaderNumber(frozenView[2], 2)); // 冻结第三行 → 3

        // 主区 RowRangeView(store,3,7)：第 0 个 RowView 是 store 行 3 → 行号 4（与冻结区 1..3 无缝衔接）
        var mainView = new RowRangeView(store, 3, 7);
        Assert.Equal(4, InvokeRowHeaderNumber(mainView[0], 0)); // 主区第一行 → 4（紧接冻结区）
        Assert.Equal(10, InvokeRowHeaderNumber(mainView[6], 6)); // 主区最后一行 → 10
    }

    [Fact]
    public void RowHeaderNumber_NonFrozen_MatchesAbsoluteRow()
    {
        var store = ColumnStore.Create(["A"], 5);
        for (var r = 0; r < 5; r++)
            store.AppendRow();
        var view = new VirtualizingSortableView(store);

        // 非冻结：主 grid 绑 VirtualizingSortableView，第 k 行 RowView.RowIndex=k → 行号 k+1
        Assert.Equal(1, InvokeRowHeaderNumber(view[0], 0));
        Assert.Equal(5, InvokeRowHeaderNumber(view[4], 4));
    }

    [Fact]
    public void RowHeaderNumber_NonRowViewItem_FallsBackToIndexPlusOne()
    {
        Assert.Equal(8, InvokeRowHeaderNumber("not a rowview", 7));
        Assert.Equal(1, InvokeRowHeaderNumber(null, 0));
    }

    // ── #7：IsTextInputFocused —— 焦点在 TextBox 时不劫持 Delete/Backspace ──
    // 正向用例（焦点=TextBox）需 STA + WPF 控件构造，本项目无 StaFact 包，改用真机 UIA 验证（见 status.md）。
    // 此处只锁定 null（无焦点）→ false 的负向契约（不需 STA）。

    [Fact]
    public void IsTextInputFocused_Null_ReturnsFalse()
    {
        Assert.False(InvokeIsTextInputFocused(null));
    }

    // ── 反射辅助 ──

    private static int InvokeRowHeaderNumber(object? rowItem, int fallbackIndex) =>
        (int)
            MainWindowType
                .GetMethod("RowHeaderNumber", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [rowItem, fallbackIndex])!;

    private static bool InvokeIsTextInputFocused(object? focused) =>
        (bool)
            MainWindowType
                .GetMethod("IsTextInputFocused", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [focused])!;

    private static Type LoadType(string fullName)
    {
        var path = Path.GetFullPath(
            Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "NumDesTools.XlsxEditor",
                "bin",
                "Debug",
                "net9.0-windows",
                "NumDesTools.XlsxEditor.dll"
            )
        );
        return Assembly.LoadFrom(path).GetType(fullName, throwOnError: true)!;
    }
}
