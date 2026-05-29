using System.Windows;
using System.Windows.Controls;
using MahApps.Metro.Controls;
using Orientation = System.Windows.Controls.Orientation;
using TextBox = System.Windows.Controls.TextBox;

namespace NumDesTools.UI
{
    public partial class InputFormularWindow : MetroWindow
    {
        public List<string> UserInputs { get; private set; }

        public InputFormularWindow(List<string> strings)
        {
            MahAppsHelper.EnsureInitialized();
            MahAppsHelper.SetExcelOwner(this);
            InitializeComponent();
            UserInputs = [];

            foreach (string str in strings)
            {
                StackPanel panel = new StackPanel
                {
                    Orientation = Orientation.Vertical,
                    Margin = new Thickness(0, 5, 0, 5),
                };

                TextBlock textBlock = new TextBlock
                {
                    Text = $"错误链接：{str}",
                    Width = 450,
                    TextWrapping = TextWrapping.Wrap,
                };
                TextBox textBox = new TextBox
                {
                    Width = 450,
                    Margin = new Thickness(0, 5, 0, 0),
                    Text = str,
                };

                panel.Children.Add(textBlock);
                panel.Children.Add(textBox);

                ItemsControl.Items.Add(panel);
            }
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (StackPanel panel in ItemsControl.Items)
            {
                TextBox textBox = (TextBox)panel.Children[1];
                UserInputs.Add(textBox.Text);
            }
            DialogResult = true;
            Close();
        }

        private void Window_EscClose(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
                Close();
        }
    }
}
