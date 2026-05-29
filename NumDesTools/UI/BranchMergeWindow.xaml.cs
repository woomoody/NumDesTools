using System.Windows;
using LibGit2Sharp;
using WpfKey = System.Windows.Input.Key;
using WpfKeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public partial class BranchMergeWindow : MahApps.Metro.Controls.MetroWindow
{
    public record CommitEntry(string Sha, string Display);

    public bool IsCherryPick => CherryRadio.IsChecked == true;

    // Cherry 模式：目标固定为当前分支
    public string? TargetBranch =>
        IsCherryPick ? _currentBranch : TargetBranchBox.SelectedItem as string;

    public string? SourceBranch => SourceBranchBox.SelectedItem as string;

    public IReadOnlyList<string> SelectedCommits =>
        CommitList.SelectedItems.Cast<CommitEntry>().Select(c => c.Sha).ToList();

    private readonly string _gitRoot;
    private string? _currentBranch;
    private bool _loading;

    public BranchMergeWindow(string gitRoot)
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
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
            _currentBranch = repo.Head.FriendlyName;
            var branches = repo
                .Branches.Where(b => !b.IsRemote)
                .Select(b => b.FriendlyName)
                .OrderBy(n => n)
                .ToList();

            _loading = true;
            TargetBranchBox.ItemsSource = branches;
            SourceBranchBox.ItemsSource = branches;

            TargetBranchBox.SelectedItem = _currentBranch;
            CurrentBranchLabel.Text = _currentBranch;
            SourceBranchBox.SelectedItem =
                branches.FirstOrDefault(b => b != _currentBranch) ?? branches.FirstOrDefault();
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
        var target = IsCherryPick ? _currentBranch : TargetBranchBox.SelectedItem as string;
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target) || source == target)
            return;

        try
        {
            using var repo = new Repository(_gitRoot);
            var sourceBranch = repo.Branches[source];
            var targetBranch = repo.Branches[target];
            if (sourceBranch == null || targetBranch == null)
                return;

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
            TargetRow.Visibility = Visibility.Collapsed;
            CurrentBranchRow.Visibility = Visibility.Visible;
            CherryCommitRow.Visibility = Visibility.Visible;
            OkButton.Content = "开始 Cherry-pick";
            LoadCommits();
        }
        else
        {
            TargetRow.Visibility = Visibility.Visible;
            CurrentBranchRow.Visibility = Visibility.Collapsed;
            CherryCommitRow.Visibility = Visibility.Collapsed;
            OkButton.Content = "开始合并";
            StatusText.Text = string.Empty;
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
