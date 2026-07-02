using System.Text.RegularExpressions;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace NumDesTools;

// a.xlsx ↔ b.xlsx 按 Key 列跨表同步，插入位置复用 LTEData.cs 的分组算法：
// 按 Key 前缀找同组最后一行插入其后，保持同类数据连续，而不是简单追加到表尾。
internal static class XlsxCrossSync
{
    private const int HeaderRow = 2;
    private const int DataStartRow = 5;
    private const string ConfigKey = "XlsxSyncMappings";

    internal sealed record Mapping(
        string Name,
        string SourcePath,
        string SourceSheet,
        string TargetPath,
        string TargetSheet,
        string KeyColumn,
        int GroupPrefixLen,
        List<string> ForwardColumns,
        List<string> ReverseColumns
    );

    internal static List<Mapping> LoadMappings()
    {
        var json = AppServices.GlobalValue.Value.GetValueOrDefault(ConfigKey, "");
        if (string.IsNullOrWhiteSpace(json))
            return [];
        try
        {
            return JsonConvert.DeserializeObject<List<Mapping>>(json) ?? [];
        }
        catch (JsonException ex)
        {
            PluginLog.Write($"[XlsxCrossSync] 映射配置解析失败: {ex.Message}");
            return [];
        }
    }

    internal static void SaveMappings(List<Mapping> mappings) =>
        AppServices.GlobalValue.SaveValue(ConfigKey, JsonConvert.SerializeObject(mappings));

    internal static void OpenSettings() => new XlsxSyncSettingsForm(LoadMappings()).Show();

    internal static void RunForward() => RunSync(reverse: false);

    internal static void RunReverse() => RunSync(reverse: true);

    private static void RunSync(bool reverse)
    {
        var mappings = LoadMappings();
        if (mappings.Count == 0)
        {
            MessageBox.Show("还没有配置同步映射，请先点「同步设置」。", "跨表同步");
            return;
        }

        var mapping = PickMapping(mappings);
        if (mapping is null)
            return;

        try
        {
            ExecuteWithPreview(mapping, reverse);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"同步失败：{ex.Message}", "跨表同步");
        }
    }

    private static Mapping? PickMapping(List<Mapping> mappings)
    {
        if (mappings.Count == 1)
            return mappings[0];

        using var picker = new Form
        {
            Text = "选择同步映射",
            Width = 360,
            Height = 150,
            Padding = new Padding(12),
            StartPosition = FormStartPosition.CenterScreen,
            KeyPreview = true,
        };
        var combo = new ComboBox
        {
            Dock = DockStyle.Top,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Margin = new Padding(0, 0, 0, 10),
        };
        combo.Items.AddRange(mappings.Select(m => m.Name).ToArray());
        combo.SelectedIndex = 0;

        Mapping? result = null;
        var okButton = new System.Windows.Forms.Button
        {
            Text = "确定",
            Dock = DockStyle.Bottom,
            Height = 34,
        };
        okButton.Click += (_, _) =>
        {
            result = mappings[combo.SelectedIndex];
            picker.Close();
        };
        picker.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                picker.Close();
        };
        picker.Controls.Add(okButton);
        picker.Controls.Add(combo);
        picker.ShowDialog();
        return result;
    }

    private static void ExecuteWithPreview(Mapping mapping, bool reverse)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        var (fromPath, fromSheetName, toPath, toSheetName, syncCols) = reverse
            ? (
                mapping.TargetPath,
                mapping.TargetSheet,
                mapping.SourcePath,
                mapping.SourceSheet,
                mapping.ReverseColumns
            )
            : (
                mapping.SourcePath,
                mapping.SourceSheet,
                mapping.TargetPath,
                mapping.TargetSheet,
                mapping.ForwardColumns
            );

        if (syncCols.Count == 0)
        {
            MessageBox.Show(
                reverse ? "该映射未配置反向同步列。" : "该映射未配置正向同步列。",
                "跨表同步"
            );
            return;
        }
        if (!File.Exists(fromPath) || !File.Exists(toPath))
        {
            MessageBox.Show("源文件或目标文件不存在，请检查同步设置里的路径。", "跨表同步");
            return;
        }

        using var fromPkg = new ExcelPackage(new FileInfo(fromPath));
        using var toPkg = new ExcelPackage(new FileInfo(toPath));
        var fromSheet = fromPkg.Workbook.Worksheets[fromSheetName];
        var toSheet = toPkg.Workbook.Worksheets[toSheetName];
        if (fromSheet is null || toSheet is null)
        {
            MessageBox.Show("找不到配置的 Sheet 名称，请检查同步设置。", "跨表同步");
            return;
        }

        var (updates, inserts, deletes) = ExecuteSync(
            fromSheet,
            toSheet,
            mapping.KeyColumn,
            mapping.GroupPrefixLen,
            syncCols,
            preview: true
        );

        if (updates == 0 && inserts == 0 && deletes == 0)
        {
            MessageBox.Show("没有需要同步的差异。", "跨表同步");
            return;
        }

        var confirm = MessageBox.Show(
            $"[{mapping.Name}] {(reverse ? "反向 b→a" : "正向 a→b")}\n"
                + $"将更新 {updates} 行 / 新增 {inserts} 行 / 删除 {deletes} 行\n\n"
                + "新增行只会写入 Key 列 + 同步列，其余列留空需自行补全。\n"
                + $"确认后原地覆写 {Path.GetFileName(toPath)}（git 可回溯）。是否继续？",
            "跨表同步 - 预览确认",
            MessageBoxButtons.OKCancel
        );
        if (confirm != DialogResult.OK)
            return;

        ExecuteSync(
            fromSheet,
            toSheet,
            mapping.KeyColumn,
            mapping.GroupPrefixLen,
            syncCols,
            preview: false
        );

        try
        {
            toPkg.Save();
        }
        catch (IOException)
        {
            MessageBox.Show(
                $"{Path.GetFileName(toPath)} 当前被其他程序占用（可能在 Excel 中打开），请关闭后重试。",
                "跨表同步"
            );
            return;
        }

        MessageBox.Show($"同步完成：更新 {updates} / 新增 {inserts} / 删除 {deletes}", "跨表同步");
    }

    private static (int Updates, int Inserts, int Deletes) ExecuteSync(
        ExcelWorksheet source,
        ExcelWorksheet target,
        string keyColumnName,
        int groupPrefixLen,
        List<string> syncCols,
        bool preview
    )
    {
        var sourceKeyCol = PubMetToExcel.FindSourceCol(source, HeaderRow, keyColumnName);
        var targetKeyCol = PubMetToExcel.FindSourceCol(target, HeaderRow, keyColumnName);
        if (sourceKeyCol == -1 || targetKeyCol == -1)
            throw new InvalidOperationException(
                $"找不到 Key 列「{keyColumnName}」，请检查同步设置。"
            );

        var sourceColIdx = new Dictionary<string, int>();
        var targetColIdx = new Dictionary<string, int>();
        foreach (var col in syncCols)
        {
            var sc = PubMetToExcel.FindSourceCol(source, HeaderRow, col);
            var tc = PubMetToExcel.FindSourceCol(target, HeaderRow, col);
            if (sc == -1 || tc == -1)
                throw new InvalidOperationException($"找不到同步列「{col}」，请检查同步设置。");
            sourceColIdx[col] = sc;
            targetColIdx[col] = tc;
        }

        var targetKeyRow = new Dictionary<string, int>(StringComparer.Ordinal);
        if (target.Dimension is not null)
        {
            for (int r = DataStartRow; r <= target.Dimension.End.Row; r++)
            {
                var k = target.Cells[r, targetKeyCol].Text?.Trim();
                if (!string.IsNullOrEmpty(k))
                    targetKeyRow[k] = r;
            }
        }

        var targetEmpty = targetKeyRow.Count == 0;
        var sourceKeys = new HashSet<string>(StringComparer.Ordinal);
        var updateOps = new List<(int TargetRow, int SourceRow)>();
        var insertOps = new List<(string Key, int SourceRow)>();

        if (source.Dimension is not null)
        {
            for (int r = DataStartRow; r <= source.Dimension.End.Row; r++)
            {
                var k = source.Cells[r, sourceKeyCol].Text?.Trim();
                if (string.IsNullOrEmpty(k))
                    continue;
                sourceKeys.Add(k);
                if (targetKeyRow.TryGetValue(k, out var existingRow))
                    updateOps.Add((existingRow, r));
                else
                    insertOps.Add((k, r));
            }
        }

        var deleteRows = targetKeyRow
            .Where(kv => !sourceKeys.Contains(kv.Key))
            .Select(kv => kv.Value)
            .ToList();

        if (preview)
            return (updateOps.Count, insertOps.Count, deleteRows.Count);

        // 2a. 删除（倒序，修正后续行号）
        deleteRows.Sort((a, b) => b.CompareTo(a));
        foreach (var rowToDel in deleteRows)
        {
            target.DeleteRow(rowToDel);
            for (int i = 0; i < updateOps.Count; i++)
                if (updateOps[i].TargetRow > rowToDel)
                    updateOps[i] = (updateOps[i].TargetRow - 1, updateOps[i].SourceRow);
        }

        // 2b. 原地更新
        foreach (var (targetRow, sourceRow) in updateOps)
        foreach (var col in syncCols)
            target.Cells[targetRow, targetColIdx[col]].Value = source
                .Cells[sourceRow, sourceColIdx[col]]
                .Value;

        // 2c. 分组插入：按 Key 前缀找同组最后一行插入其后；找不到同组或目标为空则追加末尾
        var groupedInserts = new List<(int BaseRow, string Key, int SourceRow)>();
        var tailInserts = new List<(string Key, int SourceRow)>();
        foreach (var op in insertOps)
        {
            if (targetEmpty || op.Key.Length < groupPrefixLen)
            {
                tailInserts.Add(op);
                continue;
            }
            var regex = new Regex($"^{Regex.Escape(op.Key[..groupPrefixLen])}");
            var baseRow = PubMetToExcel.FindSourceRowBlur(target, targetKeyCol, regex);
            if (baseRow == -1)
                tailInserts.Add(op);
            else
                groupedInserts.Add((baseRow, op.Key, op.SourceRow));
        }
        groupedInserts.Sort((a, b) => a.BaseRow.CompareTo(b.BaseRow));

        int rowOffset = 0;
        foreach (var (baseRow, key, sourceRow) in groupedInserts)
        {
            int writeRow = baseRow + 1 + rowOffset;
            target.InsertRow(writeRow, 1);
            rowOffset++;
            target.Cells[writeRow, targetKeyCol].Value = key;
            foreach (var col in syncCols)
                target.Cells[writeRow, targetColIdx[col]].Value = source
                    .Cells[sourceRow, sourceColIdx[col]]
                    .Value;
        }

        foreach (var (key, sourceRow) in tailInserts)
        {
            int writeRow = (target.Dimension?.End.Row ?? DataStartRow - 1) + 1;
            target.Cells[writeRow, targetKeyCol].Value = key;
            foreach (var col in syncCols)
                target.Cells[writeRow, targetColIdx[col]].Value = source
                    .Cells[sourceRow, sourceColIdx[col]]
                    .Value;
        }

        return (updateOps.Count, groupedInserts.Count + tailInserts.Count, deleteRows.Count);
    }
}
