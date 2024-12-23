using System.Windows;
using Window = System.Windows.Window;


namespace NumDesTools.UI
{
    /// <summary>
    /// SearchInputWindow.xaml 的交互逻辑
    /// </summary>
    public partial class SearchInputWindow : Window
    {
        public string SearchText { get; private set; } // 存储用户输入的关键词

        public SearchInputWindow()
        {
            InitializeComponent(); // 初始化 XAML 定义的控件
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SearchText = SearchTextBox.Text; // 获取用户输入
            DialogResult = true;            // 设置对话框结果为 true
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;           // 设置对话框结果为 false
            Close();
        }
    }
}
