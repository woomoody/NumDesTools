using System.Windows;
using Wpf.Ui.Controls;
using Key = System.Windows.Input.Key;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public enum InputDialogMode
{
    Text,
    Password,
}

public partial class InputDialog : FluentWindow
{
    public string Input { get; private set; } = string.Empty;

    private InputDialog(string prompt, string title, InputDialogMode mode, string defaultValue = "")
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();
        Title = title;
        PromptText.Text = prompt;

        switch (mode)
        {
            case InputDialogMode.Text:
                InputTextBox.Visibility = Visibility.Visible;
                InputTextBox.Text = defaultValue;
                InputTextBox.Focus();
                break;
            case InputDialogMode.Password:
                InputPasswordBox.Visibility = Visibility.Visible;
                InputPasswordBox.Focus();
                break;
        }
    }

    public static string ShowText(string title, string prompt, string defaultValue = "")
    {
        var dlg = new InputDialog(prompt, title, InputDialogMode.Text, defaultValue);
        return dlg.ShowDialog() == true ? dlg.Input : string.Empty;
    }

    public static string ShowPassword(string title, string prompt)
    {
        var dlg = new InputDialog(prompt, title, InputDialogMode.Password);
        return dlg.ShowDialog() == true ? dlg.Input : string.Empty;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Input = InputTextBox.IsVisible ? InputTextBox.Text : InputPasswordBox.Password;
        DialogResult = true;
        Close();
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
            Close();
    }
}