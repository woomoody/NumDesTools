using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MahApps.Metro.Controls;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public partial class GitConflictPickerWindow : MetroWindow
{
    public string? SelectedFile { get; private set; }
    public bool SkipHash => SkipHashBox.IsChecked == true;

    public GitConflictPickerWindow(IReadOnlyList<string> files, bool skipHash)
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        SkipHashBox.IsChecked = skipHash;
        RefreshList(files, null);
        Loaded += (_, _) => FileList.Focus();
    }

    public void RefreshList(IReadOnlyList<string> files, string? keepSelected)
    {
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

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
            Close();
    }
}
