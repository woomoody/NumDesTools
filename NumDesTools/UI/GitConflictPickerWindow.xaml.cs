using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MahApps.Metro.Controls;
using NumDesTools.ConflictResolver;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;

namespace NumDesTools.UI;

public partial class GitConflictPickerWindow : MetroWindow
{
    public string? SelectedFile { get; private set; }
    public bool SkipHash => SkipHashBox.IsChecked == true;
    public string? SelectedBranch => BranchBox.SelectedItem as string;

    private string? _gitRoot;
    private List<string> _currentFiles = [];

    public GitConflictPickerWindow(IReadOnlyList<string> files, bool skipHash, string? gitRoot = null)
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        _gitRoot = gitRoot;
        BatchAutoBtn.IsEnabled = !string.IsNullOrEmpty(gitRoot);

        SkipHashBox.IsChecked = skipHash;
        LoadBranches(gitRoot);
        RefreshList(files, null);
        Loaded += (_, _) => FileList.Focus();
    }

    private void LoadBranches(string? gitRoot)
    {
        if (string.IsNullOrEmpty(gitRoot) || !Directory.Exists(gitRoot))
        {
            BranchBox.IsEnabled = false;
            return;
        }
        try
        {
            using var repo = new LibGit2Sharp.Repository(gitRoot);
            var current = repo.Head.FriendlyName;
            var branches = repo.Branches
                .Where(b => !b.IsRemote)
                .Select(b => b.FriendlyName)
                .OrderBy(n => n)
                .ToList();

            foreach (var b in branches)
                BranchBox.Items.Add(b);

            BranchBox.SelectedItem = current;
        }
        catch
        {
            BranchBox.IsEnabled = false;
        }
    }

    public void RefreshList(IReadOnlyList<string> files, string? keepSelected)
    {
        _currentFiles = files.ToList();
        Title = $"Git 冲突解决（剩余 {files.Count} 个）";

        var prevSel = keepSelected ?? (FileList.SelectedItem as string);
        FileList.Items.Clear();
        foreach (var f in files)
            FileList.Items.Add(f);

        var idx = prevSel != null ? files.ToList().IndexOf(prevSel) : -1;
        FileList.SelectedIndex = idx >= 0 ? idx : (files.Count > 0 ? 0 : -1);

        // dynamic width: measure longest item
        var maxLen = files.Count > 0 ? files.Max(f => f.Length) : 20;
        Width = Math.Max(420, Math.Min(900, maxLen * 7.5 + 48));
    }

    private void Resolve_Click(object sender, RoutedEventArgs e)
    {
        if (FileList.SelectedItem is string f)
        {
            SelectedFile = f;
            DialogResult = true;
        }
    }

    private void FileList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (FileList.SelectedItem is string f)
        {
            SelectedFile = f;
            DialogResult = true;
        }
    }

    private void FileList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && FileList.SelectedItem is string f)
        {
            SelectedFile = f;
            DialogResult = true;
            e.Handled = true;
        }
    }

    private void SkipHash_Changed(object sender, RoutedEventArgs e)
    {
        AppServices.GlobalValue.SaveValue(
            "ConflictSkipHashFiles",
            SkipHashBox.IsChecked == true ? "true" : "false"
        );
    }

    private async void BatchAuto_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_gitRoot) || _currentFiles.Count == 0)
            return;

        BatchAutoBtn.IsEnabled = false;
        ResolveBtn.IsEnabled = false;
        BatchAutoBtn.Content = "⏳ 扫描中…";

        var files = _currentFiles.ToList();
        List<string> manual = [];
        List<string> errors = [];

        var progress = new Progress<(int done, int total, string file)>(p =>
        {
            if (p.total > 0)
                BatchAutoBtn.Content = $"⏳ {p.done}/{p.total}";
        });

        try
        {
            (manual, errors) = await Task.Run(() =>
                ExcelConflictEntry.BatchAutoResolve(_gitRoot!, files, progress)
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"批量自动解决失败：{ex.Message}", "错误");
        }
        finally
        {
            BatchAutoBtn.IsEnabled = true;
            ResolveBtn.IsEnabled = true;
            BatchAutoBtn.Content = "⚡ 一键自动解决（扫描可预选文件）";
        }

        var autoCount = files.Count - manual.Count;
        RefreshList(manual, null);

        var sb = new System.Text.StringBuilder();
        if (autoCount > 0)
            sb.AppendLine($"已自动解决 {autoCount} 个文件。");
        if (manual.Count > 0)
            sb.AppendLine($"剩余 {manual.Count} 个文件需手动处理（双方均有改动）。");
        if (errors.Count > 0)
            sb.AppendLine($"\n⚠ {errors.Count} 个文件处理出错（已保留在列表）：\n{string.Join("\n", errors)}");
        if (autoCount == 0 && errors.Count == 0)
            sb.AppendLine("所有文件均有需人工判断的冲突格，无法自动解决。");

        MessageBox.Show(sb.ToString().Trim(), autoCount > 0 ? "批量自动解决完成" : "提示");
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
            Close();
    }
}
