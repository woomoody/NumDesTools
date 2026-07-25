using System.IO;
using System.Text.Json;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 冻结窗格配置：全局默认 + 单文件/单表覆盖。列冻结 = 原生 FrozenColumnCount；
/// 行冻结 = 双 DataGrid（顶部冻结行 grid + 主 grid，共享同 DataTable 的两个 LCV 视图按行号切分）。
/// 配置文件：Documents\NumDesTools\Config\xlsx-editor-freeze.json
/// </summary>
internal static class FreezeConfig
{
    private static readonly string ConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesTools",
        "Config",
        "xlsx-editor-freeze.json"
    );

    private static FreezeData _data = new();

    static FreezeConfig() => Load();

    public static void Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                _data = JsonSerializer.Deserialize<FreezeData>(json) ?? new FreezeData();
            }
        }
        catch
        {
            // 配置损坏时用默认值
        }
    }

    public static void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
            var json = JsonSerializer.Serialize(
                _data,
                new JsonSerializerOptions { WriteIndented = true }
            );
            File.WriteAllText(ConfigPath, json);
        }
        catch
        {
            // 配置写入失败不阻塞
        }
    }

    /// <summary>
    /// 获取指定文件/工作表的冻结（列数,行数），优先级：单表覆盖 > 全局默认。
    /// </summary>
    public static (int Cols, int Rows) GetFreeze(string fileName, string sheetName)
    {
        if (
            _data.Files.TryGetValue(fileName, out var sheets)
            && sheets.TryGetValue(sheetName, out var sheet)
        )
            return (sheet.FrozenColumns, sheet.FrozenRows);
        return (_data.Global.FrozenColumns, _data.Global.FrozenRows);
    }

    /// <summary>
    /// 设置指定文件/工作表的冻结（列数,行数）。
    /// </summary>
    public static void SetFreeze(string fileName, string sheetName, int cols, int rows)
    {
        if (!_data.Files.TryGetValue(fileName, out var sheets))
        {
            sheets = new Dictionary<string, SheetFreeze>();
            _data.Files[fileName] = sheets;
        }

        // 覆盖式写：保留未传的字段会丢，所以这里两个都显式赋值
        sheets[sheetName] = new SheetFreeze { FrozenColumns = cols, FrozenRows = rows };
        Save();
    }

    /// <summary>
    /// 清除指定文件/工作表的冻结配置（列+行，回退到全局默认）。
    /// </summary>
    public static void ClearFreeze(string fileName, string sheetName)
    {
        if (_data.Files.TryGetValue(fileName, out var sheets) && sheets.Remove(sheetName))
            Save();
    }

    public static (int Cols, int Rows) GlobalFreeze
    {
        get => (_data.Global.FrozenColumns, _data.Global.FrozenRows);
        set
        {
            _data.Global.FrozenColumns = value.Cols;
            _data.Global.FrozenRows = value.Rows;
            Save();
        }
    }

    private sealed class FreezeData
    {
        public GlobalFreezeData Global { get; set; } = new();
        public Dictionary<string, Dictionary<string, SheetFreeze>> Files { get; set; } = new();
    }

    private sealed class GlobalFreezeData
    {
        public int FrozenColumns { get; set; }
        public int FrozenRows { get; set; }
    }

    private sealed class SheetFreeze
    {
        public int FrozenColumns { get; set; }
        public int FrozenRows { get; set; }
    }
}
