namespace NumDesTools.XlsxEditor;

/// <summary>
/// 列式内存存储，替代 DataTable（DataTable 有 ~4.3× 开销）。
/// 每列一个 <see cref="string"/> 数组（<c>_columns[col][row]</c>），配字符串驻留池复用重复值，
/// 并做 (row,col) 级脏跟踪供增量写回。所有单元格值统一按文本存储（与既有全 string 列模型一致）。
/// 非线程安全：读写须在同一线程（配合 UI 线程或加载线程）。
/// </summary>
public sealed class ColumnStore
{
    private const int MinCapacity = 4;

    private readonly List<string> _columnNames;
    private string?[][] _columns; // _columns[col][row]
    private int _rowCount;
    private int _rowCapacity;
    private readonly Dictionary<string, string> _internPool = new(StringComparer.Ordinal);
    private readonly HashSet<(int Row, int Col)> _dirty = [];
    // 仅在首次编辑某格时记录原始值（加载后的值），供"改回原值"时取消脏标记。
    // 字典只在编辑时增长，不为全表预分配；ClearDirty 时清空。
    private readonly Dictionary<(int Row, int Col), string?> _originalValues = [];
    private bool _structureChanged;

    private ColumnStore(IReadOnlyList<string> columnNames, int initialRowCapacity)
    {
        _columnNames = [.. columnNames];
        _rowCapacity = Math.Max(initialRowCapacity, 0);
        _columns = new string?[_columnNames.Count][];
        for (var col = 0; col < _columns.Length; col++)
        {
            _columns[col] = _rowCapacity is 0 ? [] : new string?[_rowCapacity];
        }
    }

    /// <summary>
    /// 用给定列名和可选初始行容量创建空存储。<paramref name="initialRowCapacity"/>
    /// 仅预留数组空间，不产生行（<see cref="RowCount"/> 仍为 0）。
    /// </summary>
    public static ColumnStore Create(IReadOnlyList<string> columnNames, int initialRowCapacity = 0)
    {
        ArgumentNullException.ThrowIfNull(columnNames);
        ArgumentOutOfRangeException.ThrowIfNegative(initialRowCapacity);
        return new ColumnStore(columnNames, initialRowCapacity);
    }

    public int RowCount => _rowCount;

    public int ColumnCount => _columns.Length;

    public IReadOnlyList<string> ColumnNames => _columnNames;

    public IReadOnlyCollection<(int Row, int Col)> DirtyCells => _dirty;

    /// <summary>
    /// 自上次 <see cref="ClearDirty"/> 以来是否发生过结构性改动（插/删行、扩列）。
    /// 保存路径据此决定：false → 只写 <see cref="DirtyCells"/>（增量）；true → 整表全量写回
    /// （行号相对原文件已移位，无法逐格增量写）。
    /// </summary>
    public bool StructureChanged => _structureChanged;

    /// <summary>
    /// 在末尾追加一个空行，返回其行索引。加载场景的快路径：不标脏、不移位。
    /// </summary>
    public int AppendRow()
    {
        EnsureRowCapacity(_rowCount + 1);
        return _rowCount++;
    }

    public string? GetCell(int row, int col)
    {
        ValidateRow(row);
        ValidateColumn(col);
        return _columns[col][row];
    }

    /// <summary>
    /// 写入单元格并标脏。写入值走字符串驻留池，等值内容复用同一引用以省内存。
    /// 首次编辑某格时记录原始值；若新值等于原始值（改回原值），自动取消脏标记——绿框消失。
    /// </summary>
    public void SetCell(int row, int col, string? value)
    {
        ValidateRow(row);
        ValidateColumn(col);

        // 首次编辑此格：记录原始值（加载后的当前值）
        if (!_originalValues.ContainsKey((row, col)))
        {
            _originalValues[(row, col)] = _columns[col][row];
        }

        // 新值等于原始值 → 改回原值，取消脏标记
        if (string.Equals(_originalValues[(row, col)], value, StringComparison.Ordinal))
        {
            _columns[col][row] = Intern(value);
            _dirty.Remove((row, col));
            return;
        }

        _columns[col][row] = Intern(value);
        _dirty.Add((row, col));
    }

    /// <summary>
    /// 写入单元格但<b>不</b>标脏（加载/构建阶段用），同样走驻留池。
    /// </summary>
    public void SetCellQuiet(int row, int col, string? value)
    {
        ValidateRow(row);
        ValidateColumn(col);
        _columns[col][row] = Intern(value);
    }

    public bool IsDirty(int row, int col)
    {
        ValidateRow(row);
        ValidateColumn(col);
        return _dirty.Contains((row, col));
    }

    /// <summary>
    /// 在 <paramref name="at"/> 处插入空行，原有行整体下移。<paramref name="at"/>
    /// 可等于 <see cref="RowCount"/>（等价追加）。
    /// <para>
    /// P4 变更：不再清空脏跟踪，而是 <b>remap</b>——脏集合中 row &gt;= at 的条目行号 +1（随行下移），
    /// 脏标记不丢失。置 <see cref="StructureChanged"/> = true（保存路径据此走全量写回）。
    /// </para>
    /// </summary>
    public void InsertRow(int at)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(at);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(at, _rowCount);

        EnsureRowCapacity(_rowCount + 1);
        var moveCount = _rowCount - at;
        foreach (var column in _columns)
        {
            if (moveCount > 0)
            {
                Array.Copy(column, at, column, at + 1, moveCount);
            }

            column[at] = null;
        }

        _rowCount++;
        RemapDirtyForInsert(at);
        _structureChanged = true;
    }

    /// <summary>
    /// 删除 <paramref name="at"/> 行，后续行整体上移。
    /// <para>
    /// P4 变更：不再清空脏跟踪，而是 <b>remap</b>——脏集合中 row == at 的条目丢弃，
    /// row &gt; at 的行号 -1（随行上移）。置 <see cref="StructureChanged"/> = true。
    /// </para>
    /// </summary>
    public void DeleteRow(int at)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(at);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(at, _rowCount);

        var moveCount = _rowCount - at - 1;
        foreach (var column in _columns)
        {
            if (moveCount > 0)
            {
                Array.Copy(column, at + 1, column, at, moveCount);
            }

            column[_rowCount - 1] = null;
        }

        _rowCount--;
        RemapDirtyForDelete(at);
        _structureChanged = true;
    }

    /// <summary>清空脏跟踪并重置 <see cref="StructureChanged"/>。保存成功后调用。
    /// 同时清空原始值记录——保存后当前值即新基准，下次编辑重新记录。</summary>
    public void ClearDirty()
    {
        _dirty.Clear();
        _originalValues.Clear();
        _structureChanged = false;
    }

    /// <summary>
    /// 扩展列数至 <paramref name="count"/>（用于处理 Sylvan 逐行列数不齐的 jagged 数据）。
    /// 新列按现有行容量补齐空数组。不缩减。
    /// </summary>
    public void EnsureColumnCount(int count, Func<int, string> nameFactory)
    {
        ArgumentNullException.ThrowIfNull(nameFactory);
        if (count <= _columns.Length)
        {
            return;
        }

        var expanded = new string?[count][];
        Array.Copy(_columns, expanded, _columns.Length);
        for (var col = _columns.Length; col < count; col++)
        {
            expanded[col] = _rowCapacity is 0 ? new string?[_rowCount] : new string?[_rowCapacity];
            _columnNames.Add(nameFactory(col));
        }

        _columns = expanded;
        _structureChanged = true;
    }

    /// <summary>
    /// 删除最末一列（撤销"在最右插入列"用）。列数 &lt;= 0 时无操作。丢弃该列的脏条目，置
    /// <see cref="StructureChanged"/> = true。仅支持删末列（<see cref="EnsureColumnCount"/> 也只加末列，对称）。
    /// </summary>
    public void RemoveLastColumn()
    {
        if (_columns.Length is 0)
        {
            return;
        }

        var lastCol = _columns.Length - 1;
        var shrunk = new string?[lastCol][];
        Array.Copy(_columns, shrunk, lastCol);
        _columns = shrunk;
        _columnNames.RemoveAt(lastCol);

        var affected = _dirty.Where(cell => cell.Col == lastCol).ToList();
        foreach (var cell in affected)
        {
            _dirty.Remove(cell);
        }

        _structureChanged = true;
    }

    /// <summary>InsertRow 后 remap：row &gt;= at 的脏条目行号 +1。</summary>
    private void RemapDirtyForInsert(int at)
    {
        if (_dirty.Count is 0)
        {
            return;
        }

        var affected = _dirty.Where(cell => cell.Row >= at).ToList();
        foreach (var cell in affected)
        {
            _dirty.Remove(cell);
        }

        foreach (var cell in affected)
        {
            _dirty.Add((cell.Row + 1, cell.Col));
        }
    }

    /// <summary>DeleteRow 后 remap：row == at 的脏条目丢弃，row &gt; at 的行号 -1。</summary>
    private void RemapDirtyForDelete(int at)
    {
        if (_dirty.Count is 0)
        {
            return;
        }

        var affected = _dirty.Where(cell => cell.Row >= at).ToList();
        foreach (var cell in affected)
        {
            _dirty.Remove(cell);
        }

        foreach (var cell in affected.Where(cell => cell.Row > at))
        {
            _dirty.Add((cell.Row - 1, cell.Col));
        }
    }

    private string? Intern(string? value)
    {
        if (value is null)
        {
            return null;
        }

        if (_internPool.TryGetValue(value, out var existing))
        {
            return existing;
        }

        _internPool[value] = value;
        return value;
    }

    private void EnsureRowCapacity(int required)
    {
        if (required <= _rowCapacity)
        {
            return;
        }

        var newCapacity = Math.Max(MinCapacity, _rowCapacity is 0 ? required : _rowCapacity * 2);
        if (newCapacity < required)
        {
            newCapacity = required;
        }

        for (var col = 0; col < _columns.Length; col++)
        {
            var grown = new string?[newCapacity];
            Array.Copy(_columns[col], grown, _rowCount);
            _columns[col] = grown;
        }

        _rowCapacity = newCapacity;
    }

    private void ValidateRow(int row)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(row);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(row, _rowCount);
    }

    private void ValidateColumn(int col)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(col);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(col, _columns.Length);
    }
}
