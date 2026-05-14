using System.Windows;
using Wpf.Ui.Controls;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using Key = System.Windows.Input.Key;

namespace NumDesTools.UI;

public partial class InputBoxDialog : FluentWindow
{
    public string Input { get; private set; } = string.Empty;

    public InputBoxDialog(string prompt, string title)
    {
        InitializeComponent();
        Title = title;
        PromptText.Text = prompt;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Input = InputBox.Text;
        DialogResult = true;
        Close();
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
            Close();
    }
}
