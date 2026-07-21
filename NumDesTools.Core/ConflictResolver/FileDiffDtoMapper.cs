namespace NumDesTools.ConflictResolver;

/// <summary>FileDiff ↔ FileDiffDto 映射。只序列化纯数据，不碰 INotifyPropertyChanged/computed 属性。</summary>
public static class FileDiffDtoMapper
{
    public static FileDiffDto ToDto(
        this FileDiff diff,
        string? oursLabel = null,
        string? theirsLabel = null
    ) =>
        new()
        {
            OursPath = diff.OursPath,
            TheirsPath = diff.TheirsPath,
            OursLabel = oursLabel,
            TheirsLabel = theirsLabel,
            Sheets = diff.Sheets.Select(ToDto).ToList(),
        };

    private static SheetDiffDto ToDto(SheetDiff sheet) =>
        new()
        {
            SheetName = sheet.SheetName,
            AllColumns = sheet.AllColumns,
            TypeRow = sheet.TypeRow,
            LabelRow = sheet.LabelRow,
            Rows = sheet.Rows.Select(ToDto).ToList(),
        };

    private static RowConflictDto ToDto(RowConflict row) =>
        new()
        {
            SheetName = row.SheetName,
            RowKey = row.RowKey,
            DiffType = row.DiffType,
            Origin = row.Origin,
            OursRowIndex = row.OursRowIndex,
            TheirsRowIndex = row.TheirsRowIndex,
            AllColumns = row.AllColumns,
            OursFullRow = row.OursFullRow?.ToDictionary(kv => kv.Key, kv => kv.Value?.ToString()),
            TheirsFullRow = row.TheirsFullRow?.ToDictionary(
                kv => kv.Key,
                kv => kv.Value?.ToString()
            ),
            Cells = row.Cells.Select(ToDto).ToList(),
            RowChoice = row.RowChoice,
            RowChoiceExplicit = row.IsResolved,
            AiSuggestion = row.AiSuggestion,
        };

    private static CellConflictDto ToDto(CellConflict cell) =>
        new()
        {
            ColName = cell.ColName,
            OursValue = cell.OursValue?.ToString(),
            TheirsValue = cell.TheirsValue?.ToString(),
            Choice = cell.Choice,
            IsExplicit = cell.IsExplicit,
        };

    /// <summary>把 Rust TUI 回传的 selections 合并到原 FileDiff（只更新 Choice/IsExplicit，保留对象图）。</summary>
    public static void ApplySelections(this FileDiff diff, SelectionResultDto result)
    {
        if (!result.Confirmed)
            return;

        foreach (var sel in result.Selections)
        {
            var sheet = diff.Sheets.FirstOrDefault(s => s.SheetName == sel.SheetName);
            if (sheet is null)
                continue;
            var row = sheet.Rows.FirstOrDefault(r => r.RowKey == sel.RowKey);
            if (row is null)
                continue;

            if (sel.ColName is null)
            {
                // 整行（OnlyOurs/OnlyTheirs）
                row.RowChoice = sel.Choice;
            }
            else
            {
                var cell = row.Cells.FirstOrDefault(c => c.ColName == sel.ColName);
                if (cell is not null)
                {
                    cell.Choice = sel.Choice;
                    cell.IsExplicit = true;
                }
            }
        }
    }
}
