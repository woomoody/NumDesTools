namespace NumDesTools.XlsxEditor;

/// <summary>
/// 单元格编辑记录，用于撤销/重做栈。
/// OldValue 是提交前的值（列的正确 .NET 类型），NewValue 是用户输入的文本。
/// </summary>
public sealed record CellEditRecord(int Row, int Col, object? OldValue, string NewValue);
