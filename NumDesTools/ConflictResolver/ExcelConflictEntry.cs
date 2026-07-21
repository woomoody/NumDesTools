using System.Text;
using System.Threading.Tasks;
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
        var gitRoot = AppServices.Config.Git.RootPath;
        if (string.IsNullOrEmpty(gitRoot) || !Directory.Exists(gitRoot))
        {
            System.Windows.MessageBox.Show(
                "未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。",
                "提示"
            );
            return;
        }
        RunConflictManager(gitRoot, AppServices.Config.Git.SkipHashFiles);
    }

    /// <summary>
    /// Scanner --conflict-manager 入口：全功能冲突管理器，不依赖 AppServices。
    /// gitRoot 由调用方从 cwd 向上找 .git 目录后传入。
    /// </summary>
    internal static void RunConflictManager(string gitRoot, bool skipHash = true)
    {
        string? lastSelected = null;

        while (true)
        {
            // 每次循环重新读取最新冲突列表（上一次 git add 后列表会缩短）
            List<string> allXlsx;
            List<string> allHashFiles = [];
            try
            {
                using var repo = new Repository(gitRoot);
                var allConflicted = repo
                    .Index.Conflicts.Select(c => c.Ours?.Path ?? c.Theirs?.Path ?? string.Empty)
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();

                allXlsx = allConflicted
                    .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // .xll 是插件构建产物，始终自动以对方版本为准，不进 picker
                var xllFiles = allConflicted
                    .Where(p => p.EndsWith(".xll", StringComparison.OrdinalIgnoreCase))
                    .ToList();
                allHashFiles.AddRange(xllFiles.Where(p => !allHashFiles.Contains(p)));

                // 收集所有带 # 的冲突文件（不限扩展名：xlsm/txt 都要）
                if (skipHash)
                    allHashFiles.AddRange(
                        allConflicted.Where(p =>
                            Path.GetFileName(p).Contains('#') && !allHashFiles.Contains(p)
                        )
                    );
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"读取 Git 状态失败：{ex.Message}", "错误");
                break;
            }

            List<string> conflictedFiles;
            {
                // allHashFiles（# 文件 + .xll）全部以对方版本为准，不进 picker
                using var repo = new Repository(gitRoot);
                foreach (var p in allHashFiles)
                    AutoAcceptTheirs(repo, gitRoot, p);

                conflictedFiles = skipHash
                    ? allXlsx.Where(p => !Path.GetFileName(p).Contains('#')).ToList()
                    : allXlsx;
            }

            if (conflictedFiles.Count == 0)
            {
                System.Windows.MessageBox.Show("所有 xlsx 冲突已全部解决。", "完成");
                break;
            }

            // 每次都 new，避免 ShowDialog 在已关闭窗口上重复调用；始终显示 picker 确保用户可返回上层
            var picker = new GitConflictPickerWindow(conflictedFiles, skipHash, gitRoot);
            picker.RefreshList(conflictedFiles, lastSelected);
            skipHash = picker.SkipHash;
            if (picker.ShowDialog() != true)
                break;
            var chosen = picker.SelectedFile!;
            lastSelected = chosen;
            skipHash = picker.SkipHash;
            var oursBranch = picker.SelectedBranch;

            var workingFilePath = Path.Combine(
                gitRoot,
                chosen.Replace('/', Path.DirectorySeparatorChar)
            );
            var applied = ExtractAndOpen(
                gitRoot,
                chosen,
                workingFilePath,
                autoGitAdd: true,
                oursBranchHint: oursBranch
            );
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
            var wbPath = (string)AppServices.App.ActiveWorkbook.FullName;
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
        var gitRoot = AppServices.Config.Git.RootPath;
        if (string.IsNullOrEmpty(gitRoot) || !Directory.Exists(gitRoot))
        {
            System.Windows.MessageBox.Show(
                "未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。",
                "提示"
            );
            return;
        }

        AppServices.GlobalValue.Value.TryGetValue("HistoryFileLastDir", out var lastDir);

        while (true)
        {
            var pickRoot = ResolveExcelRoot(gitRoot);
            var filePicker = new NumDesTools.UI.ExcelFilePickerWindow(pickRoot);
            if (filePicker.ShowDialog() != true || filePicker.SelectedFile == null)
                return;
            AppServices.GlobalValue.SaveValue(
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
            ) => new(c.sha, $"{c.shortSha}  {c.date}  {c.author, -16}  {c.message}", c.author);

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
    /// 自动发现当前 git 仓库里处于冲突状态的 xlsx，用 Rust TUI（conflict-tui.exe）解决——
    /// 不做分支切换/merge/cherry-pick 执行，那些交给你自己的 git 客户端做，这里只管解冲突。
    /// 没装 Rust TUI（cargo build --release 过一遍）就提示，不回退 WPF（按设计只走 Rust）。
    /// </summary>
    public static void OpenConflictRustTui()
    {
        var gitRoot = AppServices.Config.Git.RootPath;
        if (string.IsNullOrEmpty(gitRoot) || !Directory.Exists(gitRoot))
        {
            System.Windows.MessageBox.Show(
                "未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。",
                "提示"
            );
            return;
        }

        var xllDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath) ?? "";
        var rustTuiPath = RustTuiLauncher.FindRustTuiExe(xllDir);
        if (rustTuiPath is null)
        {
            System.Windows.MessageBox.Show(
                "找不到 conflict-tui.exe（Rust TUI）。先在 tools/conflict-tui 下\ncargo build --release 编译一份。",
                "提示"
            );
            return;
        }

        string? lastSelected = null;
        while (true)
        {
            List<string> conflictedXlsx;
            try
            {
                using var repo = new Repository(gitRoot);
                conflictedXlsx = repo
                    .Index.Conflicts.Select(c => c.Ours?.Path ?? c.Theirs?.Path ?? string.Empty)
                    .Where(p =>
                        !string.IsNullOrEmpty(p)
                        && p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                    )
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"读取 Git 状态失败：{ex.Message}", "错误");
                return;
            }

            if (conflictedXlsx.Count == 0)
            {
                System.Windows.MessageBox.Show("所有 xlsx 冲突已全部解决。", "完成");
                return;
            }

            var picker = new GitConflictPickerWindow(conflictedXlsx, false, gitRoot);
            picker.RefreshList(conflictedXlsx, lastSelected);
            if (picker.ShowDialog() != true)
                return;
            var chosen = picker.SelectedFile!;
            lastSelected = chosen;
            var oursBranch = picker.SelectedBranch;
            var workingFilePath = Path.Combine(
                gitRoot,
                chosen.Replace('/', Path.DirectorySeparatorChar)
            );

            ConflictBlobResult? blobs;
            try
            {
                blobs = ExtractConflictBlobs(gitRoot, chosen, oursBranchHint: oursBranch);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"提取冲突版本失败：{ex.Message}", "错误");
                continue;
            }
            if (blobs == null)
            {
                System.Windows.MessageBox.Show($"在 Index 中找不到冲突条目：{chosen}", "错误");
                continue;
            }

            FileDiff diff;
            try
            {
                diff = ExcelConflictDiffer.Diff(
                    blobs.Value.OursPath,
                    blobs.Value.TheirsPath,
                    blobs.Value.BasePath
                );
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"解析文件失败：{ex.Message}", "错误");
                continue;
            }

            if (diff.TotalConflictRows == 0)
            {
                System.Windows.MessageBox.Show("两个文件内容完全一致，没有需要解决的冲突。", "无差异");
                ConflictApplier.Apply(diff, workingFilePath, gitAdd: true);
                continue;
            }

            var confirmed = RustTuiLauncher.TryResolve(
                diff,
                rustTuiPath,
                blobs.Value.OursLabel,
                blobs.Value.TheirsLabel,
                out var error,
                ownConsole: true // Excel.exe 进程没有控制台，子进程得自己开一个
            );
            if (error != null)
                System.Windows.MessageBox.Show(error, "错误");
            if (!confirmed)
                continue; // 用户取消/失败，回选择器让用户重试或换一个文件

            ConflictApplier.Apply(diff, workingFilePath, gitAdd: true);
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
            Filter = "Excel 文件|*.xlsx",
        };
        if (dlgA.ShowDialog() != DialogResult.OK)
            return;

        using var dlgB = new OpenFileDialog
        {
            Title = "选择【他的】文件（THEIRS / 对比版本）",
            Filter = "Excel 文件|*.xlsx",
            InitialDirectory = Path.GetDirectoryName(dlgA.FileName),
        };
        if (dlgB.ShowDialog() != DialogResult.OK)
            return;

        OpenWindow(dlgA.FileName, dlgB.FileName, outPath: null, autoGitAdd: false);
    }

    // ── 批量自动解决 ──────────────────────────────────────────────────────────

    /// <summary>
    /// 对列表中所有 xlsx 冲突文件尝试自动解决：
    /// 三方预选后所有 Modified 格均有默认选项 → 直接 apply + git add；
    /// 有任意格双方都改了（需人工选择）→ 保留在返回列表。
    /// </summary>
    /// <param name="gitRoot">Git 仓库根目录</param>
    /// <param name="conflictFiles">相对路径列表</param>
    /// <param name="progress">进度回调（每处理一个文件调用一次）</param>
    /// <returns>(手动处理文件列表, 报错文件+原因列表)</returns>
    public static (List<string> manual, List<string> errors) BatchAutoResolve(
        string gitRoot,
        IReadOnlyList<string> conflictFiles,
        IProgress<(int done, int total, string file)>? progress = null
    )
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff");
        Directory.CreateDirectory(tmpDir);

        // ── 阶段一：并行 Diff（纯读，MiniExcel + libgit2 读取，无写操作）──────────
        // 每线程独立 Repository 实例（libgit2 非线程安全），临时文件名含 index 防碰撞
        // ours 临时文件留到 Phase 2 Apply 之后再删（ConflictApplier.Apply 内部用 diff.OursPath）
        var diffResults = new System.Collections.Concurrent.ConcurrentDictionary<string, FileDiff>(
            StringComparer.Ordinal
        );
        var diffErrors = new System.Collections.Concurrent.ConcurrentBag<(
            string relPath,
            string msg
        )>();
        int parallelDone = 0;

        Parallel.ForEach(
            conflictFiles.Select((p, i) => (p, i)),
            new ParallelOptions
            {
                MaxDegreeOfParallelism = Math.Max(1, Environment.ProcessorCount - 1),
            },
            item =>
            {
                var (relPath, idx) = item;
                var normPath = relPath.Replace('\\', '/');
                var fileName = Path.GetFileName(relPath);
                // index 后缀防止不同路径同名文件碰撞
                var oursPath = Path.Combine(tmpDir, $"batch_{idx}_ours_{fileName}");
                var theirsPath = Path.Combine(tmpDir, $"batch_{idx}_theirs_{fileName}");
                var basePath = Path.Combine(tmpDir, $"batch_{idx}_base_{fileName}");

                try
                {
                    string? resolvedBase = null;
                    // 每线程独立 Repository 实例，blob 提取后立即释放
                    using (var repo = new Repository(gitRoot))
                    {
                        var conflict = repo.Index.Conflicts[normPath];
                        if (conflict == null)
                            return; // 冲突已被其他操作解决，跳过

                        void WriteBlob(IndexEntry? entry, string outFile)
                        {
                            if (entry == null)
                                return;
                            var blob = repo.Lookup<Blob>(entry.Id);
                            using var src = blob.GetContentStream();
                            using var dst = new FileStream(
                                outFile,
                                FileMode.Create,
                                FileAccess.Write
                            );
                            src.CopyTo(dst);
                        }

                        WriteBlob(conflict.Ours, oursPath);
                        WriteBlob(conflict.Theirs, theirsPath);
                        if (conflict.Ancestor != null)
                        {
                            try
                            {
                                WriteBlob(conflict.Ancestor, basePath);
                                resolvedBase = basePath;
                            }
                            catch { }
                        }
                    }

                    // Diff 阶段（MiniExcel 读取，CPU 密集，可安全并行）
                    var diff = ExcelConflictDiffer.Diff(oursPath, theirsPath, resolvedBase);
                    diffResults[relPath] = diff;
                }
                catch (Exception ex)
                {
                    diffErrors.Add((relPath, ex.Message));
                    PluginLog.Write($"[BatchAutoResolve] Diff 失败 {fileName}: {ex}");
                }
                finally
                {
                    // theirs/base 读完即可删；ours 留到 Phase 2 Apply 之后再删
                    TryDelete(theirsPath);
                    TryDelete(basePath);
                    System.Threading.Interlocked.Increment(ref parallelDone);
                    progress?.Report((parallelDone, conflictFiles.Count, relPath));
                }
            }
        );

        // ── 阶段二：串行 Apply + git-add（Index.Write 文件级互斥，必须串行）──────
        var manual = new List<string>();
        var errors = diffErrors.Select(e => $"{Path.GetFileName(e.relPath)}: {e.msg}").ToList();
        int done = 0;

        foreach (var relPath in conflictFiles)
        {
            // Diff 阶段已报错的文件，直接列入手动处理
            if (diffErrors.Any(e => e.relPath == relPath))
            {
                manual.Add(relPath);
                done++;
                continue;
            }

            if (!diffResults.TryGetValue(relPath, out var diff))
            {
                // 未在 diffResults 中 = 冲突已解决或无需处理，跳过
                done++;
                continue;
            }

            bool allResolved = diff.Sheets.All(s =>
                s.Rows.Where(r => r.DiffType == RowDiffType.Modified).All(r => r.IsResolved)
            );

            if (!allResolved)
            {
                manual.Add(relPath);
                done++;
                continue;
            }

            try
            {
                var workingPath = Path.Combine(
                    gitRoot,
                    relPath.Replace('/', Path.DirectorySeparatorChar)
                );
                ConflictApplier.Apply(diff, workingPath, gitAdd: true);
            }
            catch (Exception ex)
            {
                manual.Add(relPath);
                errors.Add($"{Path.GetFileName(relPath)}: {ex.Message}");
                PluginLog.Write($"[BatchAutoResolve] Apply 失败 {Path.GetFileName(relPath)}: {ex}");
            }
            finally
            {
                // Apply 用完 ours 临时文件后才删
                TryDelete(diff.OursPath);
            }
            done++;
        }

        progress?.Report((done, conflictFiles.Count, string.Empty));

        if (errors.Count > 0)
            PluginLog.Write(
                $"[BatchAutoResolve] {errors.Count} 个文件处理报错:\n" + string.Join("\n", errors)
            );

        return (manual, errors);
    }

    // ── 内部 ─────────────────────────────────────────────────────────────────

    // 文件名含 # 或全部 sheet 含 #：直接用对方版本覆盖工作区并 git add
    internal static void AutoAcceptTheirs(Repository repo, string gitRoot, string relativePath)
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

            var sha8 = theirsSha.Length >= 8 ? theirsSha[..8] : theirsSha;
            PluginLog.Verbose(
                $"[SkipHash] 跳过 # 文件，以对方版本为准: {Path.GetFileName(relativePath)}  (theirs={sha8})"
            );
            // 同步追加到 merge 消息文件，与手动解决冲突保持一致
            try
            {
                ConflictApplier.AppendMergeMsgPublic(gitRoot, Path.GetFileName(relativePath));
            }
            catch { }
        }
        catch (Exception ex)
        {
            PluginLog.Write(
                $"[SkipHash] {Path.GetFileName(relativePath)} 自动接受对方版本失败: {ex.Message}"
            );
        }
    }

    /// <summary>从 git index 提取的冲突三方文件路径 + UI 展示用的分支标签，WPF/TUI 两条路径共用。</summary>
    internal readonly record struct ConflictBlobResult(
        string OursPath,
        string TheirsPath,
        string? BasePath,
        string? OursLabel,
        string? TheirsLabel,
        string? HeadBranch
    );

    // knownTheirsSha：cherry-pick --no-commit 不写 CHERRY_PICK_HEAD，调用方直接传 commit SHA
    // oursBranchHint / theirsBranchHint：调用方已知的分支名，省去反向推导
    // 返回 null=Index 中找不到该冲突条目；不依赖任何 UI，供 WPF（ExtractAndOpen）和 TUI（ConflictManagerTui）共用
    internal static ConflictBlobResult? ExtractConflictBlobs(
        string gitRoot,
        string relativePath,
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
        using var repo = new Repository(gitRoot);
        var conflict = repo.Index.Conflicts[normPath];
        if (conflict == null)
            return null;

        string? oursLabel,
            theirsLabel;
        var gitDir = repo.Info.Path;
        // rebase 下 stage 语义反转：conflict.Ours(stage2)=upstream 基线，
        // conflict.Theirs(stage3)=本地重放。两侧均存在时对调 blob 提取目标 + label，
        // 让 UI"我的=本地、他的=upstream"与 merge 直觉一致；三方合并内容对称，
        // 以"我的"(本地) 为基写回结果仍正确。
        bool isRebase =
            File.Exists(Path.Combine(gitDir, "REBASE_HEAD"))
            || Directory.Exists(Path.Combine(gitDir, "rebase-merge"));
        bool canSwap = isRebase && conflict.Ours != null && conflict.Theirs != null;

        if (canSwap)
        {
            var rebaseHeadSha = ReadGitHeadFile(gitDir, "REBASE_HEAD");
            var ontoRef = ReadRebaseFile(gitDir, "onto");
            var headName = ReadRebaseFile(gitDir, "head-name");
            var localBranch = headName?.Replace("refs/heads/", string.Empty) ?? "(nobranch)";

            var mySha8 = rebaseHeadSha != null ? ShortSha(rebaseHeadSha) : "?";
            oursLabel = $"{localBranch}  ({mySha8})  rebase";

            string? theirSha = null;
            var theirDesc = "(upstream)";
            if (ontoRef != null)
            {
                var ontoCommit = repo.Lookup<Commit>(ontoRef) ?? repo.Branches[ontoRef]?.Tip;
                if (ontoCommit != null)
                {
                    theirSha = ontoCommit.Sha;
                    var ontoBranch = repo
                        .Branches.Where(b => !b.IsRemote && b.Tip?.Sha == ontoCommit.Sha)
                        .Select(b => b.FriendlyName)
                        .FirstOrDefault();
                    theirDesc = ontoBranch ?? ShortSha(ontoCommit.Sha);
                }
            }
            theirsLabel = $"{theirDesc}  ({(theirSha != null ? ShortSha(theirSha) : "?")})";
        }
        else
        {
            try
            {
                var oursBranch = oursBranchHint ?? repo.Head.FriendlyName;
                var oursSha = repo.Head.Tip?.Sha[..8] ?? "?";
                oursLabel = $"{oursBranch}  ({oursSha})";

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

        if (canSwap)
        {
            // rebase 对调：我的(oursPath) ← 本地重放(stage3=Theirs)，
            // 他的(theirsPath) ← upstream 基线(stage2=Ours)
            WriteBlob(conflict.Theirs!, oursPath);
            WriteBlob(conflict.Ours!, theirsPath);
        }
        else
        {
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

        // rebase 下工作区分支用 head-name（branchB）使"我的"标 [当前]；否则用 HEAD 分支名
        var headBranchForUi = canSwap
            ? ReadRebaseFile(gitDir, "head-name")?.Replace("refs/heads/", string.Empty)
            : repo.Head.FriendlyName;

        return new ConflictBlobResult(
            oursPath,
            theirsPath,
            resolvedBasePath,
            oursLabel,
            theirsLabel,
            headBranchForUi
        );
    }

    /// 补做 git add：OpenWindow/ConflictTui 返回"无差异/已解决"时，冲突仍在 Index → 补一次 add。
    internal static void FinishGitAdd(string gitRoot, string relativePath)
    {
        try
        {
            using var repo = new Repository(gitRoot);
            var normPath = relativePath.Replace('\\', '/');
            if (repo.Index.Conflicts[normPath] != null)
            {
                repo.Index.Add(normPath);
                repo.Index.Write();
            }
        }
        catch { }
    }

    // 返回 true=已应用/完成，false=用户取消
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
        ConflictBlobResult? blobs;
        try
        {
            blobs = ExtractConflictBlobs(
                gitRoot,
                relativePath,
                knownTheirsSha,
                oursBranchHint,
                theirsBranchHint
            );
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"提取冲突版本失败：{ex.Message}", "错误");
            return false;
        }

        if (blobs == null)
        {
            System.Windows.MessageBox.Show($"在 Index 中找不到冲突条目：{relativePath}", "错误");
            return false;
        }

        var result = OpenWindow(
            blobs.Value.OursPath,
            blobs.Value.TheirsPath,
            outPath: workingFilePath,
            autoGitAdd: autoGitAdd,
            basePath: blobs.Value.BasePath,
            oursLabel: blobs.Value.OursLabel,
            theirsLabel: blobs.Value.TheirsLabel,
            headBranch: blobs.Value.HeadBranch
        );

        // "无差异"时 OpenWindow 返回 true 但不做 git add，冲突仍在 Index → 补做
        if (result && autoGitAdd)
            FinishGitAdd(gitRoot, relativePath);

        return result;
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

    // 读 .git/rebase-merge/ 下的状态文件（onto / head-name / orig-head 等）
    private static string? ReadRebaseFile(string gitDir, string name) =>
        File.Exists(Path.Combine(gitDir, "rebase-merge", name))
            ? File.ReadAllText(Path.Combine(gitDir, "rebase-merge", name)).Trim()
            : null;

    private static string ShortSha(string sha) => sha.Length >= 8 ? sha[..8] : sha;

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
            repo.Lookup<Commit>(sha)
            ?? throw new InvalidOperationException($"找不到提交：{sha[..8]}");
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

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch { }
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
                StandardErrorEncoding = Encoding.UTF8,
            },
        };
        proc.Start();
        // 同时读取 stdout/stderr，避免缓冲区满导致死锁
        var stdout = proc.StandardOutput.ReadToEnd();
        var stderr = proc.StandardError.ReadToEnd();
        var exited = proc.WaitForExit(30_000);
        if (!exited)
        {
            proc.Kill();
            throw new InvalidOperationException(
                $"git {arguments[..Math.Min(40, arguments.Length)]}… 超时（30s）"
            );
        }
        if (proc.ExitCode != 0)
            throw new InvalidOperationException($"git 返回 {proc.ExitCode}：{stderr.Trim()}");
        return stdout;
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
                SafeClose(waitWin);
            }
        })
        {
            IsBackground = true,
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

    // 安全派发：Dispatcher 已关闭时静默忽略，catch 内异常不逃逸到后台线程
    private static void SafeClose(System.Windows.Window win)
    {
        try
        {
            if (!win.Dispatcher.HasShutdownStarted)
                win.Dispatcher.BeginInvoke(
                    (System.Action)(
                        () =>
                        {
                            try
                            {
                                win.Close();
                            }
                            catch { }
                        }
                    )
                );
        }
        catch { }
    }
}
