using System.Windows;
using Window = System.Windows.Window;

namespace NumDesTools.UI;

public partial class PasswordDialog : Window
{
    public string Password { get; private set; } = string.Empty;

    public PasswordDialog(string prompt)
    {
        WpfUiHelper.ApplyDarkTheme(this);
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
