using System.Collections.Generic;
using System.Windows;
using Window = System.Windows.Window;

namespace NumDesTools.UI;

public partial class CloneHistoryNameWindow : Window
{
    public enum DialogResult { Confirm, Skip, Cancel }

    public DialogResult Result       { get; private set; } = DialogResult.Cancel;
    public string       GroupId      { get; private set; } = "";
    public string       InstanceName { get; private set; } = "";

    public CloneHistoryNameWindow(IEnumerable<string> existingGroups,
                                   string suggestedGroup    = "",
                                   string suggestedInstance = "")
    {
        InitializeComponent();

        foreach (var g in existingGroups)
            GroupCombo.Items.Add(g);

        GroupCombo.Text    = suggestedGroup;
        InstanceBox.Text   = suggestedInstance;
        InstanceBox.Focus();
        InstanceBox.SelectAll();
    }

    private void GroupCombo_SelectionChanged(object sender,
        System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (GroupCombo.SelectedItem is string s)
            GroupCombo.Text = s;
    }

    private void ClearGroup_Click(object sender, RoutedEventArgs e)
        => GroupCombo.Text = "";

    private void Confirm_Click(object sender, RoutedEventArgs e)
    {
        GroupId      = GroupCombo.Text.Trim();
        InstanceName = InstanceBox.Text.Trim();
        Result       = DialogResult.Confirm;
        Close();
    }

    private void Skip_Click(object sender, RoutedEventArgs e)
    {
        Result = DialogResult.Skip;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        Result = DialogResult.Cancel;
        Close();
    }

    private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape) { Result = DialogResult.Cancel; Close(); }
        if (e.Key == System.Windows.Input.Key.Enter)  { Confirm_Click(sender, e); }
    }
}
