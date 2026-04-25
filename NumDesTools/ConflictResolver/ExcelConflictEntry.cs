using LibGit2Sharp;
using NumDesTools.UI;
using Button   = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using Font     = System.Drawing.Font;
using ListBox  = System.Windows.Forms.ListBox;

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
            MessageBox.Show("未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。",
                "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 持久化选择窗口，循环处理直到无冲突或用户关闭
        Form?     picker    = null;
        ListBox?  lb        = null;
        CheckBox? skipHashCb = null;

        while (true)
        {
            // 每次循环重新读取最新冲突列表（上一次 git add 后列表会缩短）
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
                MessageBox.Show($"读取 Git 状态失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                break;
            }

            // 是否跳过 # 文件（以对方为准自动解决）
            bool skipHash = skipHashCb?.Checked
                ?? NumDesAddIn.GlobalValue.Value["ConflictSkipHashFiles"] == "true";

            List<string> conflictedFiles;
            if (skipHash)
            {
                // # 文件自动接受 Theirs
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
                picker?.Close();
                MessageBox.Show("所有 xlsx 冲突已全部解决。",
                    "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                break;
            }

            string chosen;
            if (conflictedFiles.Count == 1 && picker == null)
            {
                // 只剩一个且还没打开过选择窗口，直接处理
                chosen = conflictedFiles[0];
            }
            else
            {
                // 复用选择窗口（首次创建，后续刷新列表）
                if (picker == null || picker.IsDisposed)
                {
                    picker = new Form
                    {
                        Text            = "选择要解决的冲突文件",
                        StartPosition   = FormStartPosition.CenterScreen,
                        FormBorderStyle = FormBorderStyle.Sizable,
                        MaximizeBox     = false,
                        MinimumSize     = new System.Drawing.Size(400, 260),
                        BackColor       = System.Drawing.Color.FromArgb(30, 30, 30),
                        ForeColor       = System.Drawing.Color.FromArgb(220, 220, 220),
                    };

                    lb = new ListBox
                    {
                        Dock                = DockStyle.Fill,
                        Font                = new Font("Consolas", 10),
                        BackColor           = System.Drawing.Color.FromArgb(37, 37, 40),
                        ForeColor           = System.Drawing.Color.FromArgb(220, 220, 220),
                        BorderStyle         = BorderStyle.None,
                        SelectionMode       = SelectionMode.One,
                        HorizontalScrollbar = true,
                        IntegralHeight      = false,
                    };

                    var bottomPanel = new Panel
                    {
                        Dock      = DockStyle.Bottom,
                        Height    = 72,
                        BackColor = System.Drawing.Color.FromArgb(37, 37, 40),
                        Padding   = new Padding(10, 8, 10, 8),
                    };

                    skipHashCb = new CheckBox
                    {
                        Text      = "跳过 # 文件（以对方为准自动解决）",
                        Dock      = DockStyle.Top,
                        Height    = 26,
                        Checked   = NumDesAddIn.GlobalValue.Value["ConflictSkipHashFiles"] == "true",
                        Font      = new Font("微软雅黑", 9f),
                        ForeColor = System.Drawing.Color.FromArgb(180, 180, 180),
                    };
                    skipHashCb.CheckedChanged += (_, _) =>
                        NumDesAddIn.GlobalValue.SaveValue("ConflictSkipHashFiles",
                            skipHashCb.Checked ? "true" : "false");

                    var btn = new Button
                    {
                        Text      = "解决选中文件",
                        Dock      = DockStyle.Bottom,
                        Height    = 34,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                        ForeColor = System.Drawing.Color.White,
                        Font      = new Font("微软雅黑", 10f, System.Drawing.FontStyle.Bold),
                        Cursor    = Cursors.Hand,
                    };
                    btn.FlatAppearance.BorderSize = 0;

                    btn.Click += (_, _) => { if (lb.SelectedItem != null) picker.DialogResult = DialogResult.OK; };
                    bottomPanel.Controls.Add(skipHashCb);
                    bottomPanel.Controls.Add(btn);

                    picker.Controls.Add(lb);
                    picker.Controls.Add(bottomPanel);
                }

                // 刷新列表，保留上次选中项
                var prevSelected = lb!.SelectedItem?.ToString();
                lb.Items.Clear();
                lb.Items.AddRange(conflictedFiles.Cast<object>().ToArray());
                var idx = prevSelected != null ? conflictedFiles.IndexOf(prevSelected) : -1;
                lb.SelectedIndex = idx >= 0 ? idx : 0;

                picker.Text = $"Git 冲突解决（剩余 {conflictedFiles.Count} 个）";

                // 根据最长条目自动调整宽度（上限 900，下限 420）
                using var g = lb.CreateGraphics();
                var maxTextW = conflictedFiles
                    .Select(f => (int)g.MeasureString(f, lb.Font).Width)
                    .DefaultIfEmpty(0).Max();
                var listH = Math.Min(conflictedFiles.Count * lb.ItemHeight + 4, 400);
                picker.ClientSize = new System.Drawing.Size(
                    Math.Max(420, Math.Min(900, maxTextW + 32)),
                    listH + 72 + 8  // 72 = bottomPanel height
                );

                if (picker.ShowDialog() != DialogResult.OK) break; // 用户点了关闭
                chosen = lb.SelectedItem!.ToString()!;
            }

            var workingFilePath = Path.Combine(gitRoot, chosen.Replace('/', Path.DirectorySeparatorChar));
            ExtractAndOpen(gitRoot, chosen, workingFilePath, autoGitAdd: true);
        }

        picker?.Dispose();
    }

    /// <summary>
    /// 让用户分别选择两个 xlsx 文件，打开对比/合并窗口。
    /// 写回时弹另存为对话框（不执行 git add）。
    /// </summary>
    public static void OpenManualCompare()
    {
        using var dlgA = new OpenFileDialog
        {
            Title  = "选择【我的】文件（OURS / 基础版本）",
            Filter = "Excel 文件|*.xlsx"
        };
        if (dlgA.ShowDialog() != DialogResult.OK) return;

        using var dlgB = new OpenFileDialog
        {
            Title  = "选择【他的】文件（THEIRS / 对比版本）",
            Filter = "Excel 文件|*.xlsx",
            InitialDirectory = Path.GetDirectoryName(dlgA.FileName)
        };
        if (dlgB.ShowDialog() != DialogResult.OK) return;

        OpenWindow(dlgA.FileName, dlgB.FileName, outPath: null, autoGitAdd: false);
    }

    // ── 内部 ─────────────────────────────────────────────────────────────────

    // 文件名含 # 或全部 sheet 含 #：直接用 MERGE_HEAD（对方）版本覆盖工作区并 git add
    private static void AutoAcceptTheirs(Repository repo, string gitRoot, string relativePath)
    {
        try
        {
            var workingPath = Path.Combine(gitRoot, relativePath.Replace('/', Path.DirectorySeparatorChar));
            var commit = repo.Lookup<Commit>("MERGE_HEAD");
            if (commit == null) return;
            var entry = commit[relativePath.Replace('\\', '/')];
            if (entry == null) return;

            var blob = (Blob)entry.Target;
            Directory.CreateDirectory(Path.GetDirectoryName(workingPath)!);
            using (var src = blob.GetContentStream())
            using (var dst = new FileStream(workingPath, FileMode.Create, FileAccess.Write))
                src.CopyTo(dst);

            repo.Index.Add(relativePath.Replace('\\', '/'));
            repo.Index.Write();
        }
        catch { /* 单个文件失败不中断整体流程 */ }
    }

    private static void ExtractAndOpen(string gitRoot, string relativePath,
                                        string workingFilePath, bool autoGitAdd)
    {
        var tmpDir  = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff");
        Directory.CreateDirectory(tmpDir);

        var oursPath   = Path.Combine(tmpDir, "ours_"   + Path.GetFileName(relativePath));
        var theirsPath = Path.Combine(tmpDir, "theirs_" + Path.GetFileName(relativePath));

        try
        {
            GitShow(gitRoot, "ORIG_HEAD",  relativePath, oursPath);
            GitShow(gitRoot, "MERGE_HEAD", relativePath, theirsPath);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"提取 Git 版本失败：{ex.Message}\n\n" +
                            "请确认当前处于 merge 冲突状态（ORIG_HEAD 和 MERGE_HEAD 都存在）。",
                "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        OpenWindow(oursPath, theirsPath, outPath: workingFilePath, autoGitAdd: autoGitAdd);
    }

    private static void GitShow(string gitRoot, string rev, string relativePath, string outFile)
    {
        using var repo   = new Repository(gitRoot);
        var commit = repo.Lookup<Commit>(rev)
                     ?? throw new InvalidOperationException($"找不到 {rev} 提交");
        var entry  = commit[relativePath.Replace('\\', '/')]
                     ?? throw new InvalidOperationException($"{rev} 中找不到文件：{relativePath}");
        var blob   = (Blob)entry.Target;
        using var src = blob.GetContentStream();
        using var dst = new FileStream(outFile, FileMode.Create, FileAccess.Write);
        src.CopyTo(dst);
    }

    private static void OpenWindow(string oursPath, string theirsPath,
                                    string? outPath, bool autoGitAdd)
    {
        FileDiff diff;
        try
        {
            diff = ExcelConflictDiffer.Diff(oursPath, theirsPath);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"解析文件失败：{ex.Message}", "错误",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (diff.TotalConflictRows == 0)
        {
            MessageBox.Show("两个文件内容完全一致，没有需要解决的冲突。",
                "无差异", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // WPF 窗口需在 STA 线程上运行；ExcelDna 主线程已是 STA
        var win = new ExcelConflictWindow(diff, outPath, autoGitAdd);
        win.ShowDialog();
    }
}
