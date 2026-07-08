using System.Collections.Concurrent;
using NPOI.XSSF.UserModel;

namespace NumDesTools;

/// <summary>
/// 负责正式表（Type/Icon/Item.xlsx）按 id 查值时的本地优先 + backup 兜底。
/// 本地命中时不碰备份目录，只有本地确实没找到才按备份修改时间倒序试 backup。
/// </summary>
internal static class FormalTableLookupFallback
{
    private static readonly ConcurrentDictionary<string, string[]> BackupCandidateCache = new(
        StringComparer.OrdinalIgnoreCase
    );

    private static readonly Func<
        string,
        string,
        IReadOnlyList<string>
    > DefaultBackupCandidatesProvider = (backupRoot, liveFileName) =>
        ActivityDataBackupTool.FindAllBackups(backupRoot, liveFileName);

    internal static int FindKeyCol(
        string activeWorkbookPath,
        string targetWorkbook,
        int row,
        string searchValue,
        string targetSheet = "Sheet1",
        Func<(string BackupRoot, string LiveRoot)>? rootsProvider = null,
        Func<string, string, string?>? liveFileFinder = null,
        Func<string, string, IReadOnlyList<string>>? backupCandidatesProvider = null
    ) =>
        Lookup(
            activeWorkbookPath,
            targetWorkbook,
            workbook =>
            {
                var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
                var rowSource = sheet.GetRow(row);
                if (rowSource == null)
                    return -1;
                for (var j = rowSource.FirstCellNum; j <= rowSource.LastCellNum; j++)
                {
                    var cell = rowSource.GetCell(j);
                    if (cell?.ToString() == searchValue)
                        return j;
                }
                return -1;
            },
            -1,
            rootsProvider,
            liveFileFinder,
            backupCandidatesProvider
        );

    internal static int FindKeyRow(
        string activeWorkbookPath,
        string targetWorkbook,
        int col,
        string searchValue,
        string targetSheet = "Sheet1",
        Func<(string BackupRoot, string LiveRoot)>? rootsProvider = null,
        Func<string, string, string?>? liveFileFinder = null,
        Func<string, string, IReadOnlyList<string>>? backupCandidatesProvider = null
    ) =>
        Lookup(
            activeWorkbookPath,
            targetWorkbook,
            workbook =>
            {
                var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
                for (var i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    var rowSource = sheet.GetRow(i);
                    if (rowSource?.GetCell(col)?.ToString() == searchValue)
                        return i;
                }
                return -1;
            },
            -1,
            rootsProvider,
            liveFileFinder,
            backupCandidatesProvider
        );

    internal static string FindKeyColToRow(
        string activeWorkbookPath,
        string targetWorkbook,
        int row,
        int rowOut,
        string searchValue,
        string targetSheet = "Sheet1",
        Func<(string BackupRoot, string LiveRoot)>? rootsProvider = null,
        Func<string, string, string?>? liveFileFinder = null,
        Func<string, string, IReadOnlyList<string>>? backupCandidatesProvider = null
    ) =>
        Lookup(
            activeWorkbookPath,
            targetWorkbook,
            workbook =>
            {
                var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
                var rowSource = sheet.GetRow(row);
                if (rowSource == null)
                    return "Error未找到";
                for (var j = rowSource.FirstCellNum; j <= rowSource.LastCellNum; j++)
                {
                    if (rowSource.GetCell(j)?.ToString() == searchValue)
                        return sheet.GetRow(rowOut)?.GetCell(j)?.ToString() ?? string.Empty;
                }
                return "Error未找到";
            },
            "Error未找到",
            rootsProvider,
            liveFileFinder,
            backupCandidatesProvider
        );

    internal static string FindKeyRowToCol(
        string activeWorkbookPath,
        string targetWorkbook,
        int col,
        int outCol,
        string searchValue,
        string targetSheet = "Sheet1",
        Func<(string BackupRoot, string LiveRoot)>? rootsProvider = null,
        Func<string, string, string?>? liveFileFinder = null,
        Func<string, string, IReadOnlyList<string>>? backupCandidatesProvider = null
    ) =>
        Lookup(
            activeWorkbookPath,
            targetWorkbook,
            workbook =>
            {
                var sheet = workbook.GetSheet(targetSheet) ?? workbook.GetSheetAt(0);
                for (var i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    var rowSource = sheet.GetRow(i);
                    if (rowSource?.GetCell(col)?.ToString() == searchValue)
                        return rowSource.GetCell(outCol)?.ToString() ?? string.Empty;
                }
                return "Error未找到";
            },
            "Error未找到",
            rootsProvider,
            liveFileFinder,
            backupCandidatesProvider
        );

    internal static void ResetBackupCandidateCache() => BackupCandidateCache.Clear();

    private static TResult Lookup<TResult>(
        string activeWorkbookPath,
        string targetWorkbook,
        Func<XSSFWorkbook, TResult> workbookQuery,
        TResult notFound,
        Func<(string BackupRoot, string LiveRoot)>? rootsProvider,
        Func<string, string, string?>? liveFileFinder,
        Func<string, string, IReadOnlyList<string>>? backupCandidatesProvider
    )
    {
        rootsProvider ??= ActivityDataBackupTool.LoadRoots;
        liveFileFinder ??= ActivityDataBackupTool.FindFileUnder;
        backupCandidatesProvider ??= DefaultBackupCandidatesProvider;

        var workbookName = Path.GetFileName(targetWorkbook);
        var isTrackedTable = ActivityDataBackupTool.IsTrackedTableFile(workbookName);

        foreach (
            var candidate in BuildPrimaryCandidates(
                activeWorkbookPath,
                targetWorkbook,
                workbookName,
                isTrackedTable,
                rootsProvider,
                liveFileFinder
            )
        )
        {
            if (
                TryQueryWorkbook(candidate, workbookQuery, out var result)
                && !EqualityComparer<TResult>.Default.Equals(result, notFound)
            )
                return result;
        }

        if (!isTrackedTable)
            return notFound;

        var (backupRoot, _) = rootsProvider();
        if (string.IsNullOrWhiteSpace(backupRoot))
            return notFound;

        foreach (
            var backupPath in GetBackupCandidates(
                backupRoot,
                workbookName,
                backupCandidatesProvider
            )
        )
        {
            if (
                TryQueryWorkbook(backupPath, workbookQuery, out var result)
                && !EqualityComparer<TResult>.Default.Equals(result, notFound)
            )
                return result;
        }

        return notFound;
    }

    private static List<string> BuildPrimaryCandidates(
        string activeWorkbookPath,
        string targetWorkbook,
        string workbookName,
        bool isTrackedTable,
        Func<(string BackupRoot, string LiveRoot)> rootsProvider,
        Func<string, string, string?> liveFileFinder
    )
    {
        var candidates = new List<string>();
        if (isTrackedTable)
        {
            var (_, liveRoot) = rootsProvider();
            if (!string.IsNullOrWhiteSpace(liveRoot))
            {
                var livePath = liveFileFinder(liveRoot, workbookName);
                AddCandidate(candidates, livePath);
            }
        }

        AddCandidate(candidates, Path.Combine(activeWorkbookPath, targetWorkbook));
        return candidates;
    }

    private static IReadOnlyList<string> GetBackupCandidates(
        string backupRoot,
        string liveFileName,
        Func<string, string, IReadOnlyList<string>> backupCandidatesProvider
    )
    {
        if (!ReferenceEquals(backupCandidatesProvider, DefaultBackupCandidatesProvider))
            return backupCandidatesProvider(backupRoot, liveFileName);

        var cacheKey = $"{backupRoot}|{liveFileName}";
        return BackupCandidateCache.GetOrAdd(
            cacheKey,
            _ => backupCandidatesProvider(backupRoot, liveFileName).ToArray()
        );
    }

    private static void AddCandidate(List<string> candidates, string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate))
            return;
        if (!candidates.Contains(candidate, StringComparer.OrdinalIgnoreCase))
            candidates.Add(candidate);
    }

    private static bool TryQueryWorkbook<TResult>(
        string workbookPath,
        Func<XSSFWorkbook, TResult> workbookQuery,
        out TResult result
    )
    {
        result = default!;
        if (!File.Exists(workbookPath))
            return false;

        try
        {
            using var fs = new FileStream(
                workbookPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite
            );
            using var workbook = new XSSFWorkbook(fs);
            result = workbookQuery(workbook);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
