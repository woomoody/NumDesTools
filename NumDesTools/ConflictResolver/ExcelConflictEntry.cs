using System.Text;
using LibGit2Sharp;
using NumDesTools.UI;

namespace NumDesTools.ConflictResolver;

/// <summary>
/// Ribbon 按钮的两个入口：Git 冲突解决 + 手动双文件对比。
/// </summary>
public static class ExcelConflictEntry
{
    /// <summary>
    /// 自动检测当前 Git 仓库中处于冲突状态（UU）的 xlsx 文件，
    /// 让用户选择一个，提取 ORIG_HEAD/MERGE_HEAD 两版本，打开对比窗口。
    /// </summary>
    public static void OpenGitConflict()
    {
        var gitRoot = NumDesAddIn.GitRootPath;
        if (string.IsNullOrEmpty(gitRoot) || !Directory.Exists(gitRoot))
        {
            System.Windows.MessageBox.Show("未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。", "提示");
            return;
        }

        string? lastSelected = null;
        bool skipHash = NumDesAddIn.GlobalValue.Value["ConflictSkipHashFiles"] == "true";

        while (true)
        {
            // 每次循环重新读取最新冲突列表（上一次 git add 后列表会缩短）
            List<string> allXlsx;
            try
            {
                using var repo = new Repository(gitRoot);
                allXlsx = repo
                    .Index.Conflicts.Select(c => c.Ours?.Path ?? c.Theirs?.Path ?? string.Empty)
                    .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"读取 Git 状态失败：{ex.Message}", "错误");
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
                System.Windows.MessageBox.Show("所有 xlsx 冲突已全部解决。", "完成");
                break;
            }

            // 每次都 new，避免 ShowDialog 在已关闭窗口上重复调用；始终显示 picker 确保用户可返回上层
            var picker = new GitConflictPickerWindow(conflictedFiles, skipHash);
            picker.RefreshList(conflictedFiles, lastSelected);
            skipHash = picker.SkipHash;
            if (picker.ShowDialog() != true)
                break;
            var chosen = picker.SelectedFile!;
            lastSelected = chosen;
            skipHash = picker.SkipHash;

            var workingFilePath = Path.Combine(
                gitRoot,
                chosen.Replace('/', Path.DirectorySeparatorChar)
            );
            var applied = ExtractAndOpen(gitRoot, chosen, workingFilePath, autoGitAdd: true);
            if (!applied)
                continue;
        }
    }

    /// 从当前活动工作簿路径或 GitRootPath 推算 Excel 文件扫描根目录。
    /// 策略：沿当前工作簿路径向上找名为 "Excels" 的祖先目录；
    ///       若找不到，则沿 GitRootPath 向上找；
    ///       都找不到则用 GitRootPath 本身。
    private static string ResolveExcelRoot(string gitRoot)
    {
        // 1. 从当前活动工作簿路径推算
        try
        {
            var wbPath = (string)NumDesAddIn.App.ActiveWorkbook.FullName;
            if (!string.IsNullOrEmpty(wbPath))
            {
                var dir = Path.GetDirectoryName(wbPath);
                while (!string.IsNullOrEmpty(dir))
                {
                    if (Path.GetFileName(dir).Equals("Excels", StringComparison.OrdinalIgnoreCase))
                        return dir;
                    dir = Path.GetDirectoryName(dir);
                }
            }
        }
        catch { }

        // 2. 从 GitRootPath 向上找 Excels
        {
            var dir = gitRoot;
            while (!string.IsNullOrEmpty(dir))
            {
                if (Path.GetFileName(dir).Equals("Excels", StringComparison.OrdinalIgnoreCase))
                    return dir;
                dir = Path.GetDirectoryName(dir);
            }
        }

        // 3. 退回 GitRootPath
        return gitRoot;
    }

    /// <summary>
    /// 让用户选择一个 xlsx 文件，浏览其 Git 提交历史，
    /// 可选 "历史版本 vs 工作区"（支持写回 + git add）
    /// 或 "任意两个历史版本对比"（不写回）。
    /// 提交历史分页加载：初始 30 条，滚动到底自动追加。
    /// </summary>
    public static void OpenGitHistory()
    {
        var gitRoot = NumDesAddIn.GitRootPath;
        if (string.IsNullOrEmpty(gitRoot) || !Directory.Exists(gitRoot))
        {
            System.Windows.MessageBox.Show("未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。", "提示");
            return;
        }

        NumDesAddIn.GlobalValue.Value.TryGetValue("HistoryFileLastDir", out var lastDir);

        while (true)
        {
            var pickRoot = ResolveExcelRoot(gitRoot);
            var filePicker = new NumDesTools.UI.ExcelFilePickerWindow(pickRoot);
            if (filePicker.ShowDialog() != true || filePicker.SelectedFile == null)
                return;
            NumDesAddIn.GlobalValue.SaveValue(
                "HistoryFileLastDir",
                Path.GetDirectoryName(filePicker.SelectedFile) ?? gitRoot
            );
            lastDir = Path.GetDirectoryName(filePicker.SelectedFile) ?? gitRoot;

            var absPath = filePicker.SelectedFile;
            var relativePath = Path.GetRelativePath(gitRoot, absPath).Replace('\\', '/');
            var fileName = Path.GetFileName(absPath);

            var tmpDir = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff");
            Directory.CreateDirectory(tmpDir);

            GitHistoryPickerWindow.CommitEntry ToEntry(
                (string sha, string shortSha, string date, string author, string message) c
            ) => new(c.sha, $"{c.shortSha}  {c.date}  {c.author, -16}  {c.message}");

            List<GitHistoryPickerWindow.CommitEntry> LoadPage(int skip, int size) =>
                ReadGitLogForFile(gitRoot, relativePath, skip, size).Select(ToEntry).ToList();

            // 循环：对比窗口取消后回到历史选择器
            while (true)
            {
                var picker = new GitHistoryPickerWindow(
                    $"选择历史版本 — {fileName}",
                    LoadPage,
                    ["working", "another"]
                );
                if (picker.ShowDialog() != true)
                    break;

                var selectedSha = picker.SelectedSha!;
                var selectedShortSha = selectedSha[..Math.Min(7, selectedSha.Length)];
                var mode = picker.SelectedMode!;
                var snapshot = picker.LoadedEntries.ToList();

                if (mode == "working")
                {
                    var histPath = Path.Combine(tmpDir, $"hist_{selectedShortSha}_{fileName}");
                    try
                    {
                        GitShowBySha(gitRoot, selectedSha, relativePath, histPath);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"提取历史版本失败：{ex.Message}", "错误");
                        continue;
                    }
                    OpenWindow(
                        histPath,
                        absPath,
                        outPath: absPath,
                        autoGitAdd: true,
                        oursLabel: selectedShortSha,
                        theirsLabel: "工作区"
                    );
                }
                else
                {
                    var firstIdx = snapshot.FindIndex(e => e.Sha == selectedSha);
                    var picker2 = new GitHistoryPickerWindow(
                        $"选择第二个历史版本 — {fileName}（第一个：{selectedShortSha}）",
                        LoadPage,
                        ["ok"],
                        snapshot,
                        Math.Min(firstIdx + 1, snapshot.Count - 1)
                    );
                    if (picker2.ShowDialog() != true)
                        continue;

                    var sha2 = picker2.SelectedSha!;
                    var shortSha2 = sha2[..Math.Min(7, sha2.Length)];
                    var histPath = Path.Combine(tmpDir, $"hist_{selectedShortSha}_{fileName}");
                    var histPath2 = Path.Combine(tmpDir, $"hist_{shortSha2}_{fileName}");
                    try
                    {
                        GitShowBySha(gitRoot, selectedSha, relativePath, histPath);
                        GitShowBySha(gitRoot, sha2, relativePath, histPath2);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show($"提取历史版本失败：{ex.Message}", "错误");
                        continue;
                    }
                    OpenWindow(
                        histPath,
                        histPath2,
                        outPath: null,
                        autoGitAdd: false,
                        oursLabel: selectedShortSha,
                        theirsLabel: shortSha2
                    );
                }
            }
            // 历史窗口关闭后回到文件选择器
        } // end while(true)
    }

    /// <summary>
    /// 分支 Merge 或 Cherry-pick，冲突 xlsx 用现有冲突解决窗口处理。
    /// </summary>
    public static void OpenBranchMerge()
    {
        var gitRoot = NumDesAddIn.GitRootPath;
        if (string.IsNullOrEmpty(gitRoot) || !Directory.Exists(gitRoot))
        {
            System.Windows.MessageBox.Show("未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。", "提示");
            return;
        }

        // 检查工作区是否干净
        using (var repo = new Repository(gitRoot))
        {
            var status = repo.RetrieveStatus(new StatusOptions { IncludeUntracked = false });
            if (status.IsDirty)
            {
                var r = System.Windows.MessageBox.Show(
                    "工作区有未提交的改动，继续操作可能产生冲突。\n建议先 stash 或提交。\n\n仍要继续？",
                    "工作区不干净",
                    System.Windows.MessageBoxButton.YesNo
                );
                if (r != System.Windows.MessageBoxResult.Yes)
                    return;
            }
        }

        var picker = new BranchMergeWindow(gitRoot);
        if (picker.ShowDialog() != true)
            return;

        var target = picker.TargetBranch!;
        var source = picker.SourceBranch!;
        var isCherryPick = picker.IsCherryPick;
        var selectedCommits = picker.SelectedCommits;

        // 切换到目标分支
        string? switchError = null;
        var switchWin = new NumDesTools.UI.DiffProgressWindow("正在切换分支…", $"正在切换到 {target}…");
        switchWin.Loaded += (_, _) =>
        {
            var t = new System.Threading.Thread(() =>
            {
                try
                {
                    var cur = RunGit(gitRoot, "rev-parse --abbrev-ref HEAD").Trim();
                    if (cur != target)
                        RunGit(gitRoot, $"checkout \"{target}\"");
                }
                catch (Exception ex)
                {
                    switchError = ex.Message;
                }
                finally
                {
                    switchWin.Dispatcher.BeginInvoke((System.Action)(() => switchWin.Close()));
                }
            })
            {
                IsBackground = true,
            };
            t.Start();
        };
        switchWin.ShowDialog();

        if (switchError != null)
        {
            System.Windows.MessageBox.Show($"切换分支失败：{switchError}", "错误");
            return;
        }

        if (isCherryPick)
        {
            // 多 commit cherry-pick：逐个执行，每个单独解冲突并 commit
            for (var i = 0; i < selectedCommits.Count; i++)
            {
                var sha = selectedCommits[i];
                var sha8 = sha[..Math.Min(8, sha.Length)];
                var label = $"[{i + 1}/{selectedCommits.Count}] {sha8}";

                string? pickError = null;
                var progressWin = new NumDesTools.UI.DiffProgressWindow(
                    $"cherry-pick {label}…",
                    $"正在 cherry-pick {label}…"
                );
                progressWin.Loaded += (_, _) =>
                {
                    var t = new System.Threading.Thread(() =>
                    {
                        try
                        {
                            RunGit(gitRoot, $"cherry-pick --no-commit \"{sha}\"");
                        }
                        catch (Exception ex)
                        {
                            pickError = ex.Message;
                        }
                        finally
                        {
                            progressWin.Dispatcher.BeginInvoke(
                                (System.Action)(() => progressWin.Close())
                            );
                        }
                    })
                    {
                        IsBackground = true,
                    };
                    t.Start();
                };
                progressWin.ShowDialog();

                if (pickError != null)
                {
                    System.Windows.MessageBox.Show($"cherry-pick {label} 失败：{pickError}", "错误");
                    AbortOperation(gitRoot, isCherryPick: true);
                    return;
                }

                // 解当前 commit 的 xlsx 冲突
                var aborted = false;
                if (!ResolveXlsxConflicts(gitRoot, sha, source, target, label, ref aborted))
                {
                    if (aborted)
                        AbortOperation(gitRoot, isCherryPick: true);
                    return;
                }

                // commit 当前 cherry-pick
                CommitCherryPick(gitRoot, sha, sha8);
            }

            System.Windows.MessageBox.Show(
                $"cherry-pick 完成，共 {selectedCommits.Count} 个 commit。",
                "完成"
            );
        }
        else
        {
            // merge：一次性执行，统一解冲突后 commit
            string? mergeError = null;
            var progressWin = new NumDesTools.UI.DiffProgressWindow(
                $"正在 merge {source} → {target}…",
                $"正在 merge {source} → {target}，请稍候…"
            );
            progressWin.Loaded += (_, _) =>
            {
                var t = new System.Threading.Thread(() =>
                {
                    try
                    {
                        RunGit(gitRoot, $"merge --no-commit --no-ff \"{source}\"");
                    }
                    catch (Exception ex)
                    {
                        mergeError = ex.Message;
                    }
                    finally
                    {
                        progressWin.Dispatcher.BeginInvoke(
                            (System.Action)(() => progressWin.Close())
                        );
                    }
                })
                {
                    IsBackground = true,
                };
                t.Start();
            };
            progressWin.ShowDialog();

            if (mergeError != null)
            {
                System.Windows.MessageBox.Show($"merge 失败：{mergeError}", "错误");
                return;
            }

            var aborted = false;
            if (!ResolveXlsxConflicts(gitRoot, null, source, target, null, ref aborted))
            {
                if (aborted)
                    AbortOperation(gitRoot, isCherryPick: false);
                return;
            }

            CommitMerge(gitRoot, source);
        }
    }

    // 解当前暂存区中所有 xlsx 冲突，返回 true=全部解完可继续，false=用户中止
    // aborted 输出：用户选择了中止操作
    private static bool ResolveXlsxConflicts(
        string gitRoot,
        string? knownTheirsSha,
        string source,
        string target,
        string? progressLabel,
        ref bool aborted
    )
    {
        List<string> conflictedXlsx;
        try
        {
            using var repo = new Repository(gitRoot);
            conflictedXlsx = repo
                .Index.Conflicts.Select(c => c.Ours?.Path ?? c.Theirs?.Path ?? string.Empty)
                .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                .Distinct()
                .OrderBy(p => p)
                .ToList();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"读取冲突列表失败：{ex.Message}", "错误");
            return false;
        }

        var suffix = progressLabel != null ? $"\n（{progressLabel}）" : string.Empty;
        var isCherryPick = knownTheirsSha != null || progressLabel != null;

        if (conflictedXlsx.Count == 0)
        {
            var r = System.Windows.MessageBox.Show(
                $"没有 xlsx 冲突。{suffix}\n（非 xlsx 冲突需手动处理）\n\n是否继续 commit？",
                "无 xlsx 冲突",
                System.Windows.MessageBoxButton.YesNo
            );
            return r == System.Windows.MessageBoxResult.Yes;
        }

        string? lastSelected = null;
        while (true)
        {
            List<string> remaining;
            try
            {
                using var repo = new Repository(gitRoot);
                remaining = repo
                    .Index.Conflicts.Select(c => c.Ours?.Path ?? c.Theirs?.Path ?? string.Empty)
                    .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();
            }
            catch
            {
                break;
            }

            if (remaining.Count == 0)
                return true;

            var filePicker = new GitConflictPickerWindow(remaining, false);
            filePicker.RefreshList(remaining, lastSelected);
            if (filePicker.ShowDialog() != true)
            {
                var r = System.Windows.MessageBox.Show(
                    $"是否中止操作？{suffix}\n取消将保留当前冲突状态。",
                    "中止操作",
                    System.Windows.MessageBoxButton.YesNo
                );
                aborted = r == System.Windows.MessageBoxResult.Yes;
                return false;
            }

            var chosen = filePicker.SelectedFile!;
            lastSelected = chosen;
            var workingPath = Path.Combine(
                gitRoot,
                chosen.Replace('/', Path.DirectorySeparatorChar)
            );
            var applied = ExtractAndOpen(
                gitRoot,
                chosen,
                workingPath,
                autoGitAdd: true,
                knownTheirsSha,
                oursBranchHint: target,
                theirsBranchHint: isCherryPick ? $"{source}  cherry-pick" : source
            );
            if (!applied)
                continue;
        }

        return true;
    }

    private static void CommitCherryPick(string gitRoot, string sha, string sha8)
    {
        try
        {
            // 取原 commit message 作为提交信息
            var msg = RunGit(gitRoot, $"log -1 --format=%s \"{sha}\"").Trim();
            if (string.IsNullOrEmpty(msg))
                msg = $"cherry-pick {sha8}";
            RunGit(gitRoot, $"commit -m \"{msg.Replace("\"", "\\\"")}\"");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"commit 失败：{ex.Message}", "错误");
        }
    }

    private static void CommitMerge(string gitRoot, string source)
    {
        try
        {
            RunGit(gitRoot, $"commit -m \"Merge branch '{source}'\"");
            System.Windows.MessageBox.Show("merge commit 完成。", "成功");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"commit 失败：{ex.Message}", "错误");
        }
    }

    private static void AbortOperation(string gitRoot, bool isCherryPick)
    {
        try
        {
            RunGit(gitRoot, isCherryPick ? "cherry-pick --abort" : "merge --abort");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"abort 失败：{ex.Message}", "错误");
        }
    }

    /// <summary>
    /// 让用户分别选择两个 xlsx 文件，打开对比/合并窗口。
    /// 写回时弹另存为对话框（不执行 git add）。
    /// </summary>
    public static void OpenManualCompare()
    {
        using var dlgA = new OpenFileDialog
        {
            Title = "选择【我的】文件（OURS / 基础版本）",
            Filter = "Excel 文件|*.xlsx"
        };
        if (dlgA.ShowDialog() != DialogResult.OK)
            return;

        using var dlgB = new OpenFileDialog
        {
            Title = "选择【他的】文件（THEIRS / 对比版本）",
            Filter = "Excel 文件|*.xlsx",
            InitialDirectory = Path.GetDirectoryName(dlgA.FileName)
        };
        if (dlgB.ShowDialog() != DialogResult.OK)
            return;

        OpenWindow(dlgA.FileName, dlgB.FileName, outPath: null, autoGitAdd: false);
    }

    // ── 内部 ─────────────────────────────────────────────────────────────────

    // 文件名含 # 或全部 sheet 含 #：直接用对方版本覆盖工作区并 git add
    private static void AutoAcceptTheirs(Repository repo, string gitRoot, string relativePath)
    {
        try
        {
            var workingPath = Path.Combine(
                gitRoot,
                relativePath.Replace('/', Path.DirectorySeparatorChar)
            );
            var gitDir = repo.Info.Path;
            var theirsSha =
                ReadGitHeadFile(gitDir, "CHERRY_PICK_HEAD")
                ?? ReadGitHeadFile(gitDir, "MERGE_HEAD");
            if (theirsSha == null)
                return;
            var commit = repo.Lookup<Commit>(theirsSha);
            if (commit == null)
                return;
            var entry = commit[relativePath.Replace('\\', '/')];
            if (entry == null)
                return;

            var blob = (Blob)entry.Target;
            Directory.CreateDirectory(Path.GetDirectoryName(workingPath)!);
            using (var src = blob.GetContentStream())
            using (var dst = new FileStream(workingPath, FileMode.Create, FileAccess.Write))
                src.CopyTo(dst);

            repo.Index.Add(relativePath.Replace('\\', '/'));
            repo.Index.Write();
        }
        catch
        { /* 单个文件失败不中断整体流程 */
        }
    }

    // 返回 true=已应用/完成，false=用户取消
    // knownTheirsSha：cherry-pick --no-commit 不写 CHERRY_PICK_HEAD，调用方直接传 commit SHA
    // oursBranchHint / theirsBranchHint：调用方已知的分支名，省去反向推导
    private static bool ExtractAndOpen(
        string gitRoot,
        string relativePath,
        string workingFilePath,
        bool autoGitAdd,
        string? knownTheirsSha = null,
        string? oursBranchHint = null,
        string? theirsBranchHint = null
    )
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff");
        Directory.CreateDirectory(tmpDir);

        var normPath = relativePath.Replace('\\', '/');
        var oursPath = Path.Combine(tmpDir, "ours_" + Path.GetFileName(relativePath));
        var theirsPath = Path.Combine(tmpDir, "theirs_" + Path.GetFileName(relativePath));
        var basePath = Path.Combine(tmpDir, "base_" + Path.GetFileName(relativePath));

        // 优先从 Index conflict blob 直接提取：不依赖任何 HEAD 文件，merge 和 cherry-pick 均适用
        try
        {
            using var repo = new Repository(gitRoot);
            var conflict = repo.Index.Conflicts[normPath];
            if (conflict != null)
            {
                string? oursLabel,
                    theirsLabel;
                try
                {
                    var oursBranch = oursBranchHint ?? repo.Head.FriendlyName;
                    var oursSha = repo.Head.Tip?.Sha[..8] ?? "?";
                    oursLabel = $"{oursBranch}  ({oursSha})";

                    var gitDir = repo.Info.Path;
                    var theirsSha =
                        knownTheirsSha
                        ?? ReadGitHeadFile(gitDir, "CHERRY_PICK_HEAD")
                        ?? ReadGitHeadFile(gitDir, "MERGE_HEAD");
                    if (theirsSha != null)
                    {
                        var sha8 = theirsSha.Length >= 8 ? theirsSha[..8] : theirsSha;
                        if (theirsBranchHint != null)
                        {
                            theirsLabel = $"{theirsBranchHint}  ({sha8})";
                        }
                        else
                        {
                            var theirsBranch = repo
                                .Branches.Where(b => b.Tip?.Sha == theirsSha)
                                .OrderBy(b => b.IsRemote)
                                .Select(b => b.FriendlyName)
                                .FirstOrDefault();
                            theirsLabel = theirsBranch != null ? $"{theirsBranch}  ({sha8})" : sha8;
                        }
                    }
                    else
                    {
                        theirsLabel = theirsBranchHint;
                    }
                }
                catch
                {
                    oursLabel = oursBranchHint;
                    theirsLabel = theirsBranchHint;
                }

                void WriteBlob(IndexEntry? entry, string outFile)
                {
                    if (entry == null)
                        return;
                    var blob = repo.Lookup<Blob>(entry.Id);
                    using var src = blob.GetContentStream();
                    using var dst = new FileStream(outFile, FileMode.Create, FileAccess.Write);
                    src.CopyTo(dst);
                }

                WriteBlob(conflict.Ours, oursPath);

                // Theirs：优先用 Index conflict blob，fallback 到 knownTheirsSha / HEAD 文件
                if (conflict.Theirs != null)
                {
                    WriteBlob(conflict.Theirs, theirsPath);
                }
                else if (knownTheirsSha != null)
                {
                    GitShowBySha(gitRoot, knownTheirsSha, relativePath, theirsPath);
                }
                else
                {
                    var theirsSha =
                        ReadGitHeadFile(repo.Info.Path, "CHERRY_PICK_HEAD")
                        ?? ReadGitHeadFile(repo.Info.Path, "MERGE_HEAD");
                    if (theirsSha == null)
                        throw new InvalidOperationException("找不到 theirs 版本");
                    GitShowBySha(gitRoot, theirsSha, relativePath, theirsPath);
                }

                // merge-base（失败不影响主流程）
                string? resolvedBasePath = null;
                if (conflict.Ancestor != null)
                {
                    try
                    {
                        WriteBlob(conflict.Ancestor, basePath);
                        resolvedBasePath = basePath;
                    }
                    catch { }
                }

                var result = OpenWindow(
                    oursPath,
                    theirsPath,
                    outPath: workingFilePath,
                    autoGitAdd: autoGitAdd,
                    basePath: resolvedBasePath,
                    oursLabel: oursLabel,
                    theirsLabel: theirsLabel,
                    headBranch: repo.Head.FriendlyName
                );

                // "无差异"时 OpenWindow 返回 true 但不做 git add，冲突仍在 Index → 补做
                if (result && autoGitAdd)
                {
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

                return result;
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"提取冲突版本失败：{ex.Message}", "错误");
            return false;
        }

        System.Windows.MessageBox.Show($"在 Index 中找不到冲突条目：{relativePath}", "错误");
        return false;
    }

    // 直接读 gitDir（.git/ 路径）下的 HEAD 文件
    private static string? ReadGitHeadFile(string gitDir, string name)
    {
        var path = Path.Combine(gitDir, name);
        if (!File.Exists(path))
            return null;
        var sha = File.ReadAllText(path).Trim();
        return string.IsNullOrEmpty(sha) ? null : sha;
    }

    // 通过 gitRoot 工作区路径定位 .git/ 目录后读取 HEAD 文件
    private static string? ResolveHeadFileSha(string gitRoot, string name)
    {
        try
        {
            using var repo = new Repository(gitRoot);
            return ReadGitHeadFile(repo.Info.Path, name);
        }
        catch
        {
            return null;
        }
    }

    private static void GitShow(string gitRoot, string rev, string relativePath, string outFile)
    {
        using var repo = new Repository(gitRoot);
        var commit =
            repo.Lookup<Commit>(rev) ?? throw new InvalidOperationException($"找不到 {rev} 提交");
        var entry =
            commit[relativePath.Replace('\\', '/')]
            ?? throw new InvalidOperationException($"{rev} 中找不到文件：{relativePath}");
        var blob = (Blob)entry.Target;
        using var src = blob.GetContentStream();
        using var dst = new FileStream(outFile, FileMode.Create, FileAccess.Write);
        src.CopyTo(dst);
    }

    private static void GitShowBySha(
        string gitRoot,
        string sha,
        string relativePath,
        string outFile
    )
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outFile)!);
        using var repo = new Repository(gitRoot);
        var commit =
            repo.Lookup<Commit>(sha) ?? throw new InvalidOperationException($"找不到提交：{sha[..8]}");
        var entry =
            commit[relativePath.Replace('\\', '/')]
            ?? throw new InvalidOperationException($"提交 {sha[..8]} 中找不到文件：{relativePath}");
        var blob = (Blob)entry.Target;
        using var src = blob.GetContentStream();
        using var dst = new FileStream(outFile, FileMode.Create, FileAccess.Write);
        src.CopyTo(dst);
    }

    // git 可执行文件路径（延迟查找，缓存结果）
    private static string? _gitExe;
    private static string GitExe => _gitExe ??= FindGitExe();

    private static string FindGitExe()
    {
        // 1. PATH 中的 git
        foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(';'))
        {
            try
            {
                var p = Path.Combine(dir.Trim(), "git.exe");
                if (File.Exists(p))
                    return p;
            }
            catch { }
        }
        // 2. 常见安装位置
        var candidates = new[]
        {
            @"C:\Program Files\Git\bin\git.exe",
            @"C:\Program Files (x86)\Git\bin\git.exe",
        };
        foreach (var c in candidates)
            if (File.Exists(c))
                return c;
        // 3. SourceTree 内置 git（按已知路径搜索）
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var stGit = Path.Combine(
            appData,
            "Atlassian",
            "SourceTree",
            "git_local",
            "mingw32",
            "bin",
            "git.exe"
        );
        if (File.Exists(stGit))
            return stGit;
        return "git"; // 最后回退，让 OS 去找
    }

    private static string RunGit(string gitRoot, string arguments)
    {
        using var proc = new System.Diagnostics.Process
        {
            StartInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = GitExe,
                Arguments = arguments,
                WorkingDirectory = gitRoot,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
            }
        };
        proc.Start();
        var output = proc.StandardOutput.ReadToEnd();
        proc.WaitForExit(30_000);
        return output;
    }

    /// <summary>
    /// 读取文件的 git log，支持分页（skip/limit）。
    /// 使用 git 命令行替代 LibGit2Sharp，避免 pack 损坏崩溃且速度更快。
    /// </summary>
    private static List<(
        string sha,
        string shortSha,
        string date,
        string author,
        string message
    )> ReadGitLogForFile(string gitRoot, string relativePath, int skip = 0, int limit = 30)
    {
        // --format=<sha>|<date>|<author>|<subject>，用 | 分隔，避免空格歧义
        var args =
            $"log --follow --format=\"%H|%ai|%an|%s\" --skip={skip} --max-count={limit} -- \"{relativePath.Replace('/', '\\')}\"";
        var output = RunGit(gitRoot, args);

        var result = new List<(string, string, string, string, string)>();
        foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = line.Trim('"').Split('|', 4);
            if (parts.Length < 4)
                continue;
            var sha = parts[0].Trim();
            var dateRaw = parts[1].Trim();
            var author = parts[2].Trim();
            var subject = parts[3].Trim();
            if (sha.Length < 8)
                continue;

            // 解析 ISO 8601 日期（"2024-01-15 09:30:00 +0800"）
            var datePart = dateRaw.Length >= 16 ? dateRaw[..16] : dateRaw;
            result.Add((sha, sha[..8], datePart, author, subject));
        }
        return result;
    }

    // 返回 true=已应用，false=用户取消
    private static bool OpenWindow(
        string oursPath,
        string theirsPath,
        string? outPath,
        bool autoGitAdd,
        string? basePath = null,
        string? oursLabel = null,
        string? theirsLabel = null,
        string? headBranch = null
    )
    {
        FileDiff? diff = null;
        Exception? diffEx = null;

        var waitWin = new DiffProgressWindow();
        var thread = new System.Threading.Thread(() =>
        {
            try
            {
                diff = ExcelConflictDiffer.Diff(oursPath, theirsPath, basePath);
            }
            catch (Exception ex)
            {
                diffEx = ex;
            }
            finally
            {
                waitWin.Dispatcher.BeginInvoke((System.Action)(() => waitWin.Close()));
            }
        })
        {
            IsBackground = true
        };
        waitWin.Loaded += (_, _) => thread.Start();
        waitWin.ShowDialog();

        if (diffEx != null)
        {
            System.Windows.MessageBox.Show($"解析文件失败：{diffEx.Message}", "错误");
            return false;
        }

        if (diff!.TotalConflictRows == 0)
        {
            System.Windows.MessageBox.Show("两个文件内容完全一致，没有需要解决的冲突。", "无差异");
            return true;
        }

        var win = new ExcelConflictWindow(
            diff,
            outPath,
            autoGitAdd,
            oursLabel,
            theirsLabel,
            headBranch
        );
        return win.ShowDialog() == true;
    }
}
