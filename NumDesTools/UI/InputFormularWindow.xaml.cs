using System.Windows;
using System.Windows.Controls;
using Orientation = System.Windows.Controls.Orientation;
using TextBox = System.Windows.Controls.TextBox;

namespace NumDesTools.UI
{
    /// <summary>
    /// InputFormularWindow.xaml 的交互逻辑
    /// </summary>
    public partial class InputFormularWindow
    {
        public List<string> UserInputs { get; private set; }

        public InputFormularWindow(List<string> strings)
        {
            InitializeComponent();
            UserInputs = [];

            foreach (string str in strings)
            {
                StackPanel panel = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 5, 0, 5) };

                TextBlock textBlock = new TextBlock { Text = $"错误链接：{str}", Width = 450, TextWrapping = TextWrapping.Wrap };
                TextBox textBox = new TextBox { Width = 450, Margin = new Thickness(0, 5, 0, 0), Text = str };

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
    }
}
