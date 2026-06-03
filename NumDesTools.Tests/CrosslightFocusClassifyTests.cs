using NumDesTools;
using Xunit;
using static NumDesTools.CrosslightController;

namespace NumDesTools.Tests;

public class CrosslightFocusClassifyTests
{
    // --- Grid ---
    [Fact]
    public void EXCEL7_is_Grid() => Assert.Equal(FocusState.Grid, ClassifyFocusWindow("EXCEL7"));

    // --- Editing (冻结 overlay，不干扰编辑) ---
    [Fact]
    public void EDTBX_is_Editing() =>
        Assert.Equal(FocusState.Editing, ClassifyFocusWindow("EDTBX"));

    [Fact]
    public void NetUIHWND_is_Editing() =>
        Assert.Equal(FocusState.Editing, ClassifyFocusWindow("NetUIHWND"));

    [Fact]
    public void RICHEDIT60W_is_Editing() =>
        Assert.Equal(FocusState.Editing, ClassifyFocusWindow("RICHEDIT60W"));

    [Theory]
    [InlineData("EXCEL")]
    [InlineData("EXCEL.EXE")]
    [InlineData("EXCELToolbar")]
    public void EXCEL_prefix_is_Editing(string cls) =>
        Assert.Equal(FocusState.Editing, ClassifyFocusWindow(cls));

    // --- Other (隐藏 overlay) ---
    [Theory]
    [InlineData("HwndWrapper")]
    [InlineData("NativeHWND")]
    [InlineData("")]
    public void unknown_class_is_Other(string cls) =>
        Assert.Equal(FocusState.Other, ClassifyFocusWindow(cls));
}
