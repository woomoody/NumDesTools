namespace NumDesTools.Scanner;

// ── 校验严重程度 ──────────────────────────────────────────────────────────────
public enum Severity { Error, Warning, Info }

// ── 单条校验结果 ──────────────────────────────────────────────────────────────
public record ValidationIssue(
    Severity   Level,
    string     Layer,      // L1 / L2
    string     ExcelFile,
    int        Row,
    string     Field,
    string     Message
)
{
    public override string ToString()
        => $"[{Level,-7}][{Layer}] {ExcelFile}  Row {Row,5}  {Field,-28}  {Message}";
}

// ── 单张表的校验报告 ──────────────────────────────────────────────────────────
public class TableValidationReport
{
    public string          ExcelFile { get; init; } = "";
    public List<ValidationIssue> Issues    { get; }  = [];
    public bool  HasError => Issues.Any(i => i.Level == Severity.Error);
}

// ── 全量校验报告 ──────────────────────────────────────────────────────────────
public class ValidationReport
{
    public DateTime RunAt   { get; } = DateTime.Now;
    public List<TableValidationReport> Tables { get; } = [];

    public IEnumerable<ValidationIssue> AllIssues
        => Tables.SelectMany(t => t.Issues);

    public int ErrorCount   => AllIssues.Count(i => i.Level == Severity.Error);
    public int WarningCount => AllIssues.Count(i => i.Level == Severity.Warning);
}
