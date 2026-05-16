using System.Windows;
using MahApps.Metro.Controls;

namespace NumDesTools.UI;

public partial class PasswordDialog : MetroWindow
{
    public string Password { get; private set; } = string.Empty;

    public PasswordDialog(string prompt)
    {
        MahAppsHelper.EnsureInitialized();
        InitializeComponent();
        PromptText.Text = prompt;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Password = PasswordBox.Password;
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
