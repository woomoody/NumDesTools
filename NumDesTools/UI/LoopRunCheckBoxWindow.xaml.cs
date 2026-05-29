using System.Windows;
using MahApps.Metro.Controls;
using CheckBox = System.Windows.Controls.CheckBox;

namespace NumDesTools.UI
{
    public partial class LoopRunCheckBoxWindow : MetroWindow
    {
        private readonly List<object> _checkList;
        public List<object> SelectedList { get; private set; }

        public LoopRunCheckBoxWindow(List<object> inputCheckList)
        {
            MahAppsHelper.EnsureInitialized();
            MahAppsHelper.SetExcelOwner(this);
            InitializeComponent();
            _checkList = inputCheckList;
            CreateCheckBoxes();
        }

        private void CreateCheckBoxes()
        {
            // 动态创建复选框
            foreach (var check in _checkList)
            {
                CheckBox checkBox = new CheckBox
                {
                    Tag = check.ToString(),
                    Content = check + "倍率",
                    Margin = new Thickness(5),
                };
                CheckBoxContainer.Children.Add(checkBox);
            }
        }

        private void SelectAllCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var child in CheckBoxContainer.Children)
            {
                if (
                    child is CheckBox checkBox
                    && (string)checkBox.Tag != "全选"
                    && (string)checkBox.Tag != "反选"
                )
                {
                    checkBox.IsChecked = true;
                }
            }
        }

        private void SelectAllCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var child in CheckBoxContainer.Children)
            {
                if (
                    child is CheckBox checkBox
                    && (string)checkBox.Tag != "全选"
                    && (string)checkBox.Tag != "反选"
                )
                {
                    checkBox.IsChecked = false;
                }
            }
        }

        private void InvertSelectionCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var child in CheckBoxContainer.Children)
            {
                if (
                    child is CheckBox checkBox
                    && (string)checkBox.Tag != "全选"
                    && (string)checkBox.Tag != "反选"
                )
                {
                    checkBox.IsChecked = !checkBox.IsChecked;
                }
            }
            // 取消反选复选框的选中状态
            ((CheckBox)sender).IsChecked = false;
        }

        private void GetCurrentCheckBox_Click(object sender, RoutedEventArgs e)
        {
            List<object> selectedNumbers = [];

            foreach (var child in CheckBoxContainer.Children)
            {
                if (
                    child is CheckBox checkBox
                    && checkBox.IsChecked == true
                    && (string)checkBox.Tag != "全选"
                    && (string)checkBox.Tag != "反选"
                )
                {
                    selectedNumbers.Add(Convert.ToInt32(checkBox.Tag));
                }
            }

            SelectedList = selectedNumbers;
            Dispatcher.Invoke(Close); // 确保在 UI 线程上调用 Close 方法
        }

        private void Window_EscClose(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
                Close();
        }
    }
}
