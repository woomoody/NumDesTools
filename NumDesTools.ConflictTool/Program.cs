using System.Windows;
using LibGit2Sharp;
using NumDesTools.ConflictResolver;
using NumDesTools.UI;
using OfficeOpenXml;

namespace NumDesTools.ConflictTool;

internal static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        MahAppsHelper.EnsureInitialized();

        var gitRoot = ResolveGitRoot(args);
        if (gitRoot == null)
        {
            MessageBox.Show(
                "找不到 Git 仓库。\n用法：xlsx-conflict [git_root_path]",
                "xlsx 冲突解决",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            );
            return;
        }

        RunConflictLoop(gitRoot);
        System.Windows.Application.Current.Shutdown();
    }

    private static string? ResolveGitRoot(string[] args)
    {
        var startPath = args.Length > 0 && Directory.Exists(args[0])
            ? args[0]
            : Directory.GetCurrentDirectory();
        try
        {
            var discovered = Repository.Discover(startPath);
            if (discovered != null)
                return new Repository(discovered).Info.WorkingDirectory.TrimEnd('/', '\\');
        }
        catch { }
        return null;
    }

    private static void RunConflictLoop(string gitRoot)
    {
        string? lastSelected = null;
        bool skipHash = false;

        while (true)
        {
            List<string> allXlsx;
            try
            {
                using var repo = new Repository(gitRoot);
                allXlsx = repo.Index.Conflicts
                    .Select(c => c.Ours?.Path ?? c.Theirs?.Path ?? string.Empty)
                    .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取 Git 状态失败：{ex.Message}", "错误");
                break;
            }

            List<string> conflictedFiles;
            if (skipHash)
            {
                using var repo = new Repository(gitRoot);
                foreach (var p in allXlsx.Where(p => Path.GetFileName(p).Contains('#')))
                    AutoAcceptTheirs(repo, gitRoot, p);
                conflictedFiles = allXlsx.Where(p => !Path.GetFileName(p).Contains('#')).ToList();
            }
            else
            {
                conflictedFiles = allXlsx;
            }

            if (conflictedFiles.Count == 0)
            {
                MessageBox.Show("所有 xlsx 冲突已全部解决。", "完成");
                break;
            }

            var picker = new GitConflictPickerWindow(conflictedFiles, skipHash);
            picker.RefreshList(conflictedFiles, lastSelected);
            skipHash = picker.SkipHash;
            if (picker.ShowDialog() != true)
                break;
            var chosen = picker.SelectedFile!;
            lastSelected = chosen;
            skipHash = picker.SkipHash;

            var workingFilePath = Path.Combine(gitRoot, chosen.Replace('/', Path.DirectorySeparatorChar));
            ExtractAndOpen(gitRoot, chosen, workingFilePath);
        }
    }

    private static void AutoAcceptTheirs(IRepository repo, string gitRoot, string relPath)
    {
        var conflict = repo.Index.Conflicts[relPath];
        if (conflict?.Theirs == null)
            return;
        var blob = repo.Lookup<Blob>(conflict.Theirs.Id);
        var destPath = Path.Combine(gitRoot, relPath.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(destPath)!);
        using (var fs = File.Create(destPath))
            blob.GetContentStream().CopyTo(fs);
        repo.Index.Add(relPath);
        repo.Index.Write();
    }

    private static void ExtractAndOpen(string gitRoot, string relPath, string workingFilePath)
    {
        var normPath = relPath.Replace('\\', '/');
        var tmpDir = Path.Combine(Path.GetTempPath(), "NumDesXlsxConflict");
        Directory.CreateDirectory(tmpDir);
        var fileName   = Path.GetFileName(relPath);
        var oursPath   = Path.Combine(tmpDir, $"ours_{fileName}");
        var theirsPath = Path.Combine(tmpDir, $"theirs_{fileName}");
        var basePath   = Path.Combine(tmpDir, $"base_{fileName}");

        string oursBranchLabel, theirsBranchLabel, headBranchName;
        FileDiff diff;
        try
        {
            using var repo = new Repository(gitRoot);
            var conflict = repo.Index.Conflicts[normPath];
            if (conflict == null)
            {
                MessageBox.Show($"在 Index 中找不到冲突条目：{relPath}", "错误");
                return;
            }

            // Ours
            WriteBlobEntry(repo, conflict.Ours, oursPath);

            // Theirs：优先 Index blob，fallback 到 HEAD 文件
            if (conflict.Theirs != null)
            {
                WriteBlobEntry(repo, conflict.Theirs, theirsPath);
            }
            else
            {
                var theirsSha = ReadGitHeadFile(repo.Info.Path, "CHERRY_PICK_HEAD")
                    ?? ReadGitHeadFile(repo.Info.Path, "MERGE_HEAD")
                    ?? throw new InvalidOperationException("找不到 theirs 版本");
                GitShowBySha(repo, theirsSha, normPath, theirsPath);
            }

            // Base（失败不影响主流程，只是没有预选）
            string? resolvedBasePath = null;
            if (conflict.Ancestor != null)
            {
                try { WriteBlobEntry(repo, conflict.Ancestor, basePath); resolvedBasePath = basePath; }
                catch { }
            }

            headBranchName    = repo.Head.FriendlyName;
            var oursSha8      = conflict.Ours?.Id.Sha[..8] ?? "";
            var theirsSha8    = conflict.Theirs?.Id.Sha[..8] ?? "";
            oursBranchLabel   = $"{headBranchName}  ({oursSha8})";
            theirsBranchLabel = $"{ResolveTheirsBranch(repo, gitRoot)}  ({theirsSha8})";

            diff = ExcelConflictDiffer.Diff(oursPath, theirsPath, resolvedBasePath);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"提取/比较冲突版本失败：{ex.Message}", "错误");
            return;
        }

        var win = new ExcelConflictWindow(
            diff,
            outPath: workingFilePath,
            autoGitAdd: true,
            oursLabel: oursBranchLabel,
            theirsLabel: theirsBranchLabel,
            headBranch: headBranchName
        );
        win.ShowDialog();

        // "无差异"时 Window 不执行 git add，冲突仍在 Index → 补做
        try
        {
            using var repo2 = new Repository(gitRoot);
            if (repo2.Index.Conflicts[normPath] != null)
            {
                repo2.Index.Add(normPath);
                repo2.Index.Write();
            }
        }
        catch { }
    }

    private static void WriteBlobEntry(IRepository repo, IndexEntry entry, string destPath)
    {
        var blob = repo.Lookup<Blob>(entry.Id);
        using var src = blob.GetContentStream();
        using var dst = File.Create(destPath);
        src.CopyTo(dst);
    }

    private static void WriteBlob(IRepository repo, ObjectId? id, string destPath)
    {
        if (id == null) { File.WriteAllBytes(destPath, []); return; }
        var blob = repo.Lookup<Blob>(id);
        using var fs = File.Create(destPath);
        blob.GetContentStream().CopyTo(fs);
    }

    private static void GitShowBySha(IRepository repo, string sha, string relPath, string outFile)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outFile)!);
        var commit = repo.Lookup<Commit>(sha)
            ?? throw new InvalidOperationException($"找不到提交：{sha[..8]}");
        var entry = commit[relPath]
            ?? throw new InvalidOperationException($"提交 {sha[..8]} 中找不到：{relPath}");
        var blob = (Blob)entry.Target;
        using var src = blob.GetContentStream();
        using var dst = File.Create(outFile);
        src.CopyTo(dst);
    }

    private static string? ReadGitHeadFile(string gitDir, string name)
    {
        var path = Path.Combine(gitDir, name);
        if (!File.Exists(path)) return null;
        var sha = File.ReadAllText(path).Trim();
        return string.IsNullOrEmpty(sha) ? null : sha;
    }

    // 从 MERGE_HEAD / CHERRY_PICK_HEAD 文件读取 SHA，再反查分支名
    private static string ResolveTheirsBranch(IRepository repo, string gitRoot)
    {
        foreach (var headFile in new[] { "MERGE_HEAD", "CHERRY_PICK_HEAD" })
        {
            var path = Path.Combine(gitRoot, ".git", headFile);
            if (!File.Exists(path)) continue;
            var sha = File.ReadAllText(path).Trim();
            if (string.IsNullOrEmpty(sha)) continue;
            // 反查本地分支
            var branch = repo.Branches
                .Where(b => !b.IsRemote && b.Tip?.Sha == sha)
                .Select(b => b.FriendlyName)
                .FirstOrDefault();
            if (!string.IsNullOrEmpty(branch)) return branch;
            // 反查远端分支
            branch = repo.Branches
                .Where(b => b.IsRemote && b.Tip?.Sha == sha)
                .Select(b => b.FriendlyName)
                .FirstOrDefault();
            if (!string.IsNullOrEmpty(branch)) return branch;
            // 找不到分支名，返回短 SHA
            return sha[..Math.Min(8, sha.Length)];
        }
        return "";
    }
}
