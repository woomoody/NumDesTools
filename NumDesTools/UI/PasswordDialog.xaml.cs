using System.Windows;
using Wpf.Ui.Controls;

namespace NumDesTools.UI;

public partial class PasswordDialog : FluentWindow
{
    public string Password { get; private set; } = string.Empty;

    public PasswordDialog(string prompt)
    {
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
