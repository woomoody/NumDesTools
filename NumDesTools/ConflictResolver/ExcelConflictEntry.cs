using System.Text;
using LibGit2Sharp;
using NumDesTools.UI;
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using Font = System.Drawing.Font;
using ListBox = System.Windows.Forms.ListBox;

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
            MessageBox.Show(
                "未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。",
                "提示",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return;
        }

        // 持久化选择窗口，循环处理直到无冲突或用户关闭
        Form? picker = null;
        ListBox? lb = null;
        CheckBox? skipHashCb = null;

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
                MessageBox.Show(
                    $"读取 Git 状态失败：{ex.Message}",
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                break;
            }

            // 是否跳过 # 文件（以对方为准自动解决）
            bool skipHash =
                skipHashCb?.Checked
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
                MessageBox.Show(
                    "所有 xlsx 冲突已全部解决。",
                    "完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
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
                        Text = "选择要解决的冲突文件",
                        StartPosition = FormStartPosition.CenterScreen,
                        FormBorderStyle = FormBorderStyle.Sizable,
                        MaximizeBox = false,
                        MinimumSize = new System.Drawing.Size(400, 260),
                        BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                        ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
                    };

                    lb = new ListBox
                    {
                        Dock = DockStyle.Fill,
                        Font = new Font("Consolas", 10),
                        BackColor = System.Drawing.Color.FromArgb(37, 37, 40),
                        ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
                        BorderStyle = BorderStyle.None,
                        SelectionMode = SelectionMode.One,
                        HorizontalScrollbar = true,
                        IntegralHeight = false,
                    };

                    var bottomPanel = new Panel
                    {
                        Dock = DockStyle.Bottom,
                        Height = 72,
                        BackColor = System.Drawing.Color.FromArgb(37, 37, 40),
                        Padding = new Padding(10, 8, 10, 8),
                    };

                    skipHashCb = new CheckBox
                    {
                        Text = "跳过 # 文件（以对方为准自动解决）",
                        Dock = DockStyle.Top,
                        Height = 26,
                        Checked = NumDesAddIn.GlobalValue.Value["ConflictSkipHashFiles"] == "true",
                        Font = new Font("微软雅黑", 9f),
                        ForeColor = System.Drawing.Color.FromArgb(180, 180, 180),
                    };
                    skipHashCb.CheckedChanged += (_, _) =>
                        NumDesAddIn.GlobalValue.SaveValue(
                            "ConflictSkipHashFiles",
                            skipHashCb.Checked ? "true" : "false"
                        );

                    var btn = new Button
                    {
                        Text = "解决选中文件",
                        Dock = DockStyle.Bottom,
                        Height = 34,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                        ForeColor = System.Drawing.Color.White,
                        Font = new Font("微软雅黑", 10f, System.Drawing.FontStyle.Bold),
                        Cursor = Cursors.Hand,
                    };
                    btn.FlatAppearance.BorderSize = 0;

                    btn.Click += (_, _) =>
                    {
                        if (lb.SelectedItem != null)
                            picker.DialogResult = DialogResult.OK;
                    };
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
                    .DefaultIfEmpty(0)
                    .Max();
                var listH = Math.Min(conflictedFiles.Count * lb.ItemHeight + 4, 400);
                picker.ClientSize = new System.Drawing.Size(
                    Math.Max(420, Math.Min(900, maxTextW + 32)),
                    listH + 72 + 8 // 72 = bottomPanel height
                );

                if (picker.ShowDialog() != DialogResult.OK)
                    break; // 用户点了关闭
                chosen = lb.SelectedItem!.ToString()!;
            }

            var workingFilePath = Path.Combine(
                gitRoot,
                chosen.Replace('/', Path.DirectorySeparatorChar)
            );
            var applied = ExtractAndOpen(gitRoot, chosen, workingFilePath, autoGitAdd: true);
            if (!applied)
                continue; // 用户点了取消，返回文件选择
        }

        picker?.Dispose();
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
            MessageBox.Show(
                "未配置 GitRootPath，请在 NumDesToolsConfig.json 中设置。",
                "提示",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
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

            // 分页状态
            const int PageSize = 30;
            var allCommits =
                new List<(
                    string sha,
                    string shortSha,
                    string date,
                    string author,
                    string message
                )>();
            int loadedCount = 0;
            bool isLoading = false;
            bool hasMore = true;

            var picker = new Form
            {
                Text = $"选择历史版本 — {fileName}",
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimumSize = new System.Drawing.Size(700, 400),
                BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
            };

            var lb = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9.5f),
                BackColor = System.Drawing.Color.FromArgb(37, 37, 40),
                ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
                BorderStyle = BorderStyle.None,
                SelectionMode = SelectionMode.One,
                HorizontalScrollbar = true,
                IntegralHeight = false,
            };

            var statusLabel = new System.Windows.Forms.Label
            {
                Dock = DockStyle.Bottom,
                Height = 18,
                Text = "加载中…",
                ForeColor = System.Drawing.Color.FromArgb(130, 130, 130),
                Font = new Font("微软雅黑", 8f),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
            };

            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                BackColor = System.Drawing.Color.FromArgb(37, 37, 40),
                Padding = new Padding(10, 6, 10, 6),
            };

            var btnVsWorking = new Button
            {
                Text = "与当前工作区对比",
                Height = 32,
                Width = 160,
                Left = 10,
                Top = 6,
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                Font = new Font("微软雅黑", 9.5f, System.Drawing.FontStyle.Bold),
                Cursor = Cursors.Hand,
            };
            btnVsWorking.FlatAppearance.BorderSize = 0;

            var btnVsAnother = new Button
            {
                Text = "与另一历史版本对比",
                Height = 32,
                Width = 168,
                Left = 180,
                Top = 6,
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(60, 80, 60),
                ForeColor = System.Drawing.Color.White,
                Font = new Font("微软雅黑", 9.5f, System.Drawing.FontStyle.Bold),
                Cursor = Cursors.Hand,
            };
            btnVsAnother.FlatAppearance.BorderSize = 0;

            btnVsWorking.Click += (_, _) =>
            {
                if (lb.SelectedItem != null)
                {
                    picker.Tag = "working";
                    picker.DialogResult = DialogResult.OK;
                }
            };
            btnVsAnother.Click += (_, _) =>
            {
                if (lb.SelectedItem != null)
                {
                    picker.Tag = "another";
                    picker.DialogResult = DialogResult.OK;
                }
            };
            bottomPanel.Controls.Add(btnVsWorking);
            bottomPanel.Controls.Add(btnVsAnother);

            picker.Controls.Add(lb);
            picker.Controls.Add(statusLabel);
            picker.Controls.Add(bottomPanel);
            picker.ClientSize = new System.Drawing.Size(900, 500);
            picker.WindowState = FormWindowState.Maximized;

            // 追加一批到 ListBox
            void AppendBatch(
                IEnumerable<(
                    string sha,
                    string shortSha,
                    string date,
                    string author,
                    string message
                )> batch
            )
            {
                lb.BeginUpdate();
                foreach (var (_, shortSha, date, author, message) in batch)
                    lb.Items.Add($"{shortSha}  {date}  {author, -16}  {message}");
                lb.EndUpdate();
            }

            // 后台加载下一页（skip = 已加载数）
            void LoadNextPage()
            {
                if (isLoading || !hasMore)
                    return;
                isLoading = true;
                statusLabel.Text = "加载中…";

                int skip = loadedCount;
                System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                {
                    List<(string, string, string, string, string)> page;
                    try
                    {
                        page = ReadGitLogForFile(gitRoot, relativePath, skip, PageSize);
                    }
                    catch
                    {
                        page = [];
                    }

                    if (!picker.IsHandleCreated || picker.IsDisposed)
                    {
                        isLoading = false;
                        return;
                    }
                    picker.BeginInvoke(
                        (System.Action)(
                            () =>
                            {
                                allCommits.AddRange(page);
                                loadedCount += page.Count;
                                hasMore = page.Count == PageSize;
                                isLoading = false;

                                AppendBatch(page);
                                if (lb.Items.Count > 0 && lb.SelectedIndex < 0)
                                    lb.SelectedIndex = 0;

                                statusLabel.Text = hasMore
                                    ? $"已加载 {loadedCount} 条，滚动到底加载更多"
                                    : $"共 {loadedCount} 条，已全部加载";
                            }
                        )
                    );
                });
            }

            // 监听滚动：接近底部时加载下一页；Enter / 双击 = 与工作区对比；Esc = 关闭
            lb.KeyDown += (_, e) =>
            {
                if (
                    (e.KeyCode == Keys.Down || e.KeyCode == Keys.PageDown)
                    && lb.TopIndex + lb.Height / lb.ItemHeight >= lb.Items.Count - 3
                )
                    LoadNextPage();
                if (e.KeyCode == Keys.Enter && lb.SelectedItem != null)
                {
                    picker.Tag = "working";
                    picker.DialogResult = DialogResult.OK;
                }
                if (e.KeyCode == Keys.Escape)
                    picker.Close();
            };
            lb.MouseDoubleClick += (_, _) =>
            {
                if (lb.SelectedItem != null)
                {
                    picker.Tag = "working";
                    picker.DialogResult = DialogResult.OK;
                }
            };
            lb.MouseWheel += (_, e) =>
            {
                int delta = e.Delta > 0 ? -1 : 1;
                int next = Math.Clamp(lb.SelectedIndex + delta, 0, lb.Items.Count - 1);
                lb.SelectedIndex = next;
                lb.TopIndex = Math.Max(
                    0,
                    next - lb.Height / (lb.ItemHeight > 0 ? lb.ItemHeight : 1) / 2
                );
                if (
                    e.Delta < 0
                    && lb.TopIndex + lb.Height / (lb.ItemHeight + 1) >= lb.Items.Count - 5
                )
                    LoadNextPage();
            };
            picker.KeyPreview = true;
            picker.KeyDown += (_, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                    picker.Close();
                if (e.KeyCode == Keys.Tab)
                {
                    lb.Focus();
                    e.Handled = true;
                }
            };

            picker.Load += (_, _) => LoadNextPage();

            var tmpDir = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff");
            Directory.CreateDirectory(tmpDir);

            // 循环：对比窗口取消后回到历史选择器
            while (true)
            {
                if (picker.ShowDialog() != DialogResult.OK)
                    break;

                var selectedIdx = lb.SelectedIndex;
                var mode = picker.Tag?.ToString() ?? "working";
                picker.Tag = null; // 清除，避免复用时残留

                if (selectedIdx < 0 || selectedIdx >= allCommits.Count)
                    continue;
                var selectedSha = allCommits[selectedIdx].sha;

                if (mode == "working")
                {
                    var histPath = Path.Combine(
                        tmpDir,
                        $"hist_{allCommits[selectedIdx].shortSha}_{fileName}"
                    );
                    try
                    {
                        GitShowBySha(gitRoot, selectedSha, relativePath, histPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"提取历史版本失败：{ex.Message}",
                            "错误",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                        continue;
                    }
                    OpenWindow(histPath, absPath, outPath: absPath, autoGitAdd: true);
                    // 对比完成或取消后回到历史选择器
                }
                else
                {
                    // 第二个版本选择窗口
                    var allCommits2 = new List<(
                        string sha,
                        string shortSha,
                        string date,
                        string author,
                        string message
                    )>(allCommits);
                    int loadedCount2 = loadedCount;
                    bool hasMore2 = hasMore;
                    bool isLoading2 = false;

                    var picker2 = new Form
                    {
                        Text = $"选择第二个历史版本 — {fileName}（第一个：{allCommits[selectedIdx].shortSha}）",
                        StartPosition = FormStartPosition.CenterScreen,
                        FormBorderStyle = FormBorderStyle.Sizable,
                        MinimumSize = new System.Drawing.Size(700, 400),
                        BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                        ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
                        ClientSize = new System.Drawing.Size(900, 500),
                        WindowState = FormWindowState.Maximized,
                    };
                    var lb2 = new ListBox
                    {
                        Dock = DockStyle.Fill,
                        Font = new Font("Consolas", 9.5f),
                        BackColor = System.Drawing.Color.FromArgb(37, 37, 40),
                        ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
                        BorderStyle = BorderStyle.None,
                        SelectionMode = SelectionMode.One,
                        HorizontalScrollbar = true,
                        IntegralHeight = false,
                    };
                    var statusLabel2 = new System.Windows.Forms.Label
                    {
                        Dock = DockStyle.Bottom,
                        Height = 18,
                        Text = hasMore2
                            ? $"已加载 {loadedCount2} 条，滚动到底加载更多"
                            : $"共 {loadedCount2} 条，已全部加载",
                        ForeColor = System.Drawing.Color.FromArgb(130, 130, 130),
                        Font = new Font("微软雅黑", 8f),
                        TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                        Padding = new Padding(8, 0, 0, 0),
                        BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                    };
                    lb2.BeginUpdate();
                    foreach (var (_, shortSha, date, author, message) in allCommits2)
                        lb2.Items.Add($"{shortSha}  {date}  {author, -16}  {message}");
                    lb2.EndUpdate();
                    lb2.SelectedIndex = Math.Min(selectedIdx + 1, lb2.Items.Count - 1);

                    void LoadNextPage2()
                    {
                        if (isLoading2 || !hasMore2)
                            return;
                        isLoading2 = true;
                        statusLabel2.Text = "加载中…";
                        int skip2 = loadedCount2;
                        System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                        {
                            List<(string, string, string, string, string)> page2;
                            try
                            {
                                page2 = ReadGitLogForFile(gitRoot, relativePath, skip2, PageSize);
                            }
                            catch
                            {
                                page2 = [];
                            }
                            if (!picker2.IsHandleCreated || picker2.IsDisposed)
                            {
                                isLoading2 = false;
                                return;
                            }
                            picker2.BeginInvoke(
                                (System.Action)(
                                    () =>
                                    {
                                        allCommits2.AddRange(page2);
                                        loadedCount2 += page2.Count;
                                        hasMore2 = page2.Count == PageSize;
                                        isLoading2 = false;
                                        lb2.BeginUpdate();
                                        foreach (var (_, s, d, a, m) in page2)
                                            lb2.Items.Add($"{s}  {d}  {a, -16}  {m}");
                                        lb2.EndUpdate();
                                        statusLabel2.Text = hasMore2
                                            ? $"已加载 {loadedCount2} 条，滚动到底加载更多"
                                            : $"共 {loadedCount2} 条，已全部加载";
                                    }
                                )
                            );
                        });
                    }

                    lb2.MouseWheel += (_, e) =>
                    {
                        int delta2 = e.Delta > 0 ? -1 : 1;
                        int next2 = Math.Clamp(lb2.SelectedIndex + delta2, 0, lb2.Items.Count - 1);
                        lb2.SelectedIndex = next2;
                        lb2.TopIndex = Math.Max(
                            0,
                            next2 - lb2.Height / (lb2.ItemHeight > 0 ? lb2.ItemHeight : 1) / 2
                        );
                        if (
                            e.Delta < 0
                            && lb2.TopIndex + lb2.Height / (lb2.ItemHeight + 1)
                                >= lb2.Items.Count - 5
                        )
                            LoadNextPage2();
                    };
                    lb2.KeyDown += (_, e) =>
                    {
                        if (
                            (e.KeyCode == Keys.Down || e.KeyCode == Keys.PageDown)
                            && lb2.TopIndex + lb2.Height / lb2.ItemHeight >= lb2.Items.Count - 3
                        )
                            LoadNextPage2();
                        if (e.KeyCode == Keys.Enter && lb2.SelectedItem != null)
                            picker2.DialogResult = DialogResult.OK;
                        if (e.KeyCode == Keys.Escape)
                            picker2.Close();
                    };
                    lb2.MouseDoubleClick += (_, _) =>
                    {
                        if (lb2.SelectedItem != null)
                            picker2.DialogResult = DialogResult.OK;
                    };
                    picker2.KeyPreview = true;
                    picker2.KeyDown += (_, e) =>
                    {
                        if (e.KeyCode == Keys.Escape)
                            picker2.Close();
                        if (e.KeyCode == Keys.Tab)
                        {
                            lb2.Focus();
                            e.Handled = true;
                        }
                    };

                    var bottomPanel2 = new Panel
                    {
                        Dock = DockStyle.Bottom,
                        Height = 44,
                        BackColor = System.Drawing.Color.FromArgb(37, 37, 40)
                    };
                    var btnOk2 = new Button
                    {
                        Text = "开始对比",
                        Height = 32,
                        Width = 110,
                        Left = 10,
                        Top = 6,
                        FlatStyle = FlatStyle.Flat,
                        BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                        ForeColor = System.Drawing.Color.White,
                        Font = new Font("微软雅黑", 9.5f, System.Drawing.FontStyle.Bold),
                        Cursor = Cursors.Hand,
                    };
                    btnOk2.FlatAppearance.BorderSize = 0;
                    btnOk2.Click += (_, _) =>
                    {
                        if (lb2.SelectedItem != null)
                            picker2.DialogResult = DialogResult.OK;
                    };
                    bottomPanel2.Controls.Add(btnOk2);
                    picker2.Controls.Add(lb2);
                    picker2.Controls.Add(statusLabel2);
                    picker2.Controls.Add(bottomPanel2);

                    if (picker2.ShowDialog() != DialogResult.OK)
                    {
                        picker2.Dispose();
                        continue;
                    } // 取消第二个版本选择 → 回历史选择器
                    var selectedIdx2 = lb2.SelectedIndex;
                    picker2.Dispose();

                    var sha2 = allCommits2[selectedIdx2].sha;
                    var histPath = Path.Combine(
                        tmpDir,
                        $"hist_{allCommits[selectedIdx].shortSha}_{fileName}"
                    );
                    var histPath2 = Path.Combine(
                        tmpDir,
                        $"hist_{allCommits2[selectedIdx2].shortSha}_{fileName}"
                    );
                    try
                    {
                        GitShowBySha(gitRoot, selectedSha, relativePath, histPath);
                        GitShowBySha(gitRoot, sha2, relativePath, histPath2);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"提取历史版本失败：{ex.Message}",
                            "错误",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                        continue;
                    }
                    OpenWindow(histPath, histPath2, outPath: null, autoGitAdd: false);
                    // 对比完成或取消后回到历史选择器
                }
            }

            picker.Dispose();
            // 历史窗口关闭后回到文件选择器
        } // end while(true)
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

    // 文件名含 # 或全部 sheet 含 #：直接用 MERGE_HEAD（对方）版本覆盖工作区并 git add
    private static void AutoAcceptTheirs(Repository repo, string gitRoot, string relativePath)
    {
        try
        {
            var workingPath = Path.Combine(
                gitRoot,
                relativePath.Replace('/', Path.DirectorySeparatorChar)
            );
            var commit = repo.Lookup<Commit>("MERGE_HEAD");
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
    private static bool ExtractAndOpen(
        string gitRoot,
        string relativePath,
        string workingFilePath,
        bool autoGitAdd
    )
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "NumDesExcelDiff");
        Directory.CreateDirectory(tmpDir);

        var oursPath = Path.Combine(tmpDir, "ours_" + Path.GetFileName(relativePath));
        var theirsPath = Path.Combine(tmpDir, "theirs_" + Path.GetFileName(relativePath));

        try
        {
            GitShow(gitRoot, "ORIG_HEAD", relativePath, oursPath);
            GitShow(gitRoot, "MERGE_HEAD", relativePath, theirsPath);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"提取 Git 版本失败：{ex.Message}\n\n" + "请确认当前处于 merge 冲突状态（ORIG_HEAD 和 MERGE_HEAD 都存在）。",
                "错误",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return false;
        }

        return OpenWindow(oursPath, theirsPath, outPath: workingFilePath, autoGitAdd: autoGitAdd);
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
        bool autoGitAdd
    )
    {
        // 大文件 Diff 可能耗时数秒，放到后台线程并显示等待框，避免 UI 卡死
        FileDiff? diff = null;
        Exception? diffEx = null;

        using var waitForm = new Form
        {
            Text = "正在解析差异…",
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            Size = new System.Drawing.Size(320, 80),
            ControlBox = false,
            BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
        };
        waitForm.Controls.Add(
            new System.Windows.Forms.Label
            {
                Text = "正在比较文件，请稍候…",
                ForeColor = System.Drawing.Color.White,
                AutoSize = true,
                Left = 20,
                Top = 20,
            }
        );

        var thread = new System.Threading.Thread(() =>
        {
            try
            {
                diff = ExcelConflictDiffer.Diff(oursPath, theirsPath);
            }
            catch (Exception ex)
            {
                diffEx = ex;
            }
            finally
            {
                // 用 BeginInvoke 避免：若 ShowDialog 尚未建立 Handle，Invoke 会抛异常
                try
                {
                    waitForm.BeginInvoke((System.Action)waitForm.Close);
                }
                catch
                { /* 窗口已销毁，忽略 */
                }
            }
        })
        {
            IsBackground = true
        };

        // 在 Load（Handle 已建立）后再启动线程，保证 BeginInvoke 安全
        waitForm.Load += (_, _) => thread.Start();
        waitForm.ShowDialog();

        if (diffEx != null)
        {
            MessageBox.Show(
                $"解析文件失败：{diffEx.Message}",
                "错误",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return false;
        }

        if (diff!.TotalConflictRows == 0)
        {
            MessageBox.Show(
                "两个文件内容完全一致，没有需要解决的冲突。",
                "无差异",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
            return true;
        }

        var win = new ExcelConflictWindow(diff, outPath, autoGitAdd);
        return win.ShowDialog() == true;
    }
}
