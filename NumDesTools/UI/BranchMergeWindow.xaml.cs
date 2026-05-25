using System.Windows;
using LibGit2Sharp;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public partial class BranchMergeWindow : MahApps.Metro.Controls.MetroWindow
{
    public record CommitEntry(string Sha, string Display);

    public bool IsCherryPick => CherryRadio.IsChecked == true;
    public string? TargetBranch => TargetBranchBox.SelectedItem as string;
    public string? SourceBranch => SourceBranchBox.SelectedItem as string;
    public IReadOnlyList<string> SelectedCommits =>
        CommitList.SelectedItems.Cast<CommitEntry>().Select(c => c.Sha).ToList();

    private readonly string _gitRoot;
    private bool _loading;

    public BranchMergeWindow(string gitRoot)
    {
        MahAppsHelper.EnsureInitialized();
        InitializeComponent();
        _gitRoot = gitRoot;
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        LoadBranches();
    }

    private void LoadBranches()
    {
        try
        {
            using var repo = new Repository(_gitRoot);
            var current = repo.Head.FriendlyName;
            var branches = repo
                .Branches.Where(b => !b.IsRemote)
                .Select(b => b.FriendlyName)
                .OrderBy(n => n)
                .ToList();

            _loading = true;
            TargetBranchBox.ItemsSource = branches;
            SourceBranchBox.ItemsSource = branches;

            TargetBranchBox.SelectedItem = current;
            // 来源默认选非当前分支的第一个
            SourceBranchBox.SelectedItem =
                branches.FirstOrDefault(b => b != current) ?? branches.FirstOrDefault();
            _loading = false;

            if (IsCherryPick)
                LoadCommits();
        }
        catch (Exception ex)
        {
            StatusText.Text = $"读取分支失败：{ex.Message}";
        }
    }

    private void LoadCommits()
    {
        CommitList.Items.Clear();
        var source = SourceBranchBox.SelectedItem as string;
        var target = TargetBranchBox.SelectedItem as string;
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target) || source == target)
            return;

        try
        {
            using var repo = new Repository(_gitRoot);
            var sourceBranch = repo.Branches[source];
            var targetBranch = repo.Branches[target];
            if (sourceBranch == null || targetBranch == null)
                return;

            // 列出 source 比 target 多出的 commits
            var mergeBase = repo.ObjectDatabase.FindMergeBase(sourceBranch.Tip, targetBranch.Tip);
            var commits = repo
                .Commits.QueryBy(
                    new CommitFilter
                    {
                        IncludeReachableFrom = sourceBranch.Tip,
                        ExcludeReachableFrom = mergeBase,
                        SortBy = CommitSortStrategies.Topological,
                    }
                )
                .Take(200)
                .ToList();

            foreach (var c in commits)
            {
                var display =
                    $"{c.Sha[..8]}  {c.Author.When:MM-dd HH:mm}  {c.Author.Name, -14}  {c.MessageShort}";
                CommitList.Items.Add(new CommitEntry(c.Sha, display));
            }

            StatusText.Text = $"共 {commits.Count} 个可摘取的 commit";
        }
        catch (Exception ex)
        {
            StatusText.Text = $"加载 commit 失败：{ex.Message}";
        }
    }

    private void Mode_Changed(object sender, RoutedEventArgs e)
    {
        if (MergeSourceRow == null)
            return;

        if (IsCherryPick)
        {
            MergeSourceRow.Visibility = Visibility.Visible; // 来源分支两种模式都需要
            CherryCommitRow.Visibility = Visibility.Visible;
            OkButton.Content = "开始 Cherry-pick";
            LoadCommits();
        }
        else
        {
            MergeSourceRow.Visibility = Visibility.Visible;
            CherryCommitRow.Visibility = Visibility.Collapsed;
            OkButton.Content = "开始合并";
        }
    }

    private void TargetBranch_Changed(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e
    )
    {
        if (_loading)
            return;
        if (IsCherryPick)
            LoadCommits();
    }

    private void SourceBranch_Changed(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e
    )
    {
        if (_loading)
            return;
        if (IsCherryPick)
            LoadCommits();
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(TargetBranch))
        {
            StatusText.Text = "请选择目标分支";
            return;
        }
        if (string.IsNullOrEmpty(SourceBranch))
        {
            StatusText.Text = "请选择来源分支";
            return;
        }
        if (TargetBranch == SourceBranch)
        {
            StatusText.Text = "目标分支和来源分支不能相同";
            return;
        }
        if (IsCherryPick && SelectedCommits.Count == 0)
        {
            StatusText.Text = "请至少选择一个 commit";
            return;
        }
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;

    private void Window_KeyDown(object sender, WpfKeyEventArgs e)
    {
        if (e.Key == WpfKey.Escape)
            DialogResult = false;
    }
}
