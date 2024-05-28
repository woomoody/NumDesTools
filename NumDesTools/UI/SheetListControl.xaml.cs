using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using ListBox = System.Windows.Controls.ListBox;
using MenuItem = System.Windows.Controls.MenuItem;
using MessageBox = System.Windows.MessageBox;
using Style = System.Windows.Style;

namespace NumDesTools.UI
{
    /// <summary>
    /// SheetListControl.xaml 的交互逻辑
    /// </summary>
    public partial class SheetListControl
    {
        public static Application ExcelApp = NumDesAddIn.App;
        public ObservableCollection<SelfComSheetCollect> Sheets { get; } = new ObservableCollection<SelfComSheetCollect>();
        public SheetListControl()
        {
            InitializeComponent();
            var worksheets = ExcelApp.ActiveWorkbook.Sheets.Cast<Worksheet>()
                .Select(x => new SelfComSheetCollect
                {
                    Name = x.Name,
                    IsHidden = x.Visible == XlSheetVisibility.xlSheetHidden,
                    DetailInfo = (x.Cells[1, 2] as Range)?.Value2?.ToString(),
                    UsedRangeSize = new Tuple<int, int>(x.UsedRange.Rows.Count, x.UsedRange.Columns.Count)
                });

            foreach (var worksheet in worksheets)
            {
                Sheets.Add(worksheet);
            }
            ListBoxSheet.ItemsSource = Sheets;
            ListBoxSheet.DisplayMemberPath = "Name";

            SetListBoxItemStyle(ListBoxSheet);

        }

        private void ListBoxSheet_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;
            while (source is { } and not ListBoxItem)
            {
                source = VisualTreeHelper.GetParent(source);
            }

            if (source is not ListBoxItem listBoxItem)
            {
                return;
            }

            listBoxItem.IsSelected = true;

            var contextMenu = new ContextMenu();
            var showItem = new MenuItem { Header = "显示" };
            showItem.Click += (_, _) =>
            {
                foreach (SelfComSheetCollect item in ListBoxSheet.SelectedItems)
                {
                    var sheet = ExcelApp.ActiveWorkbook.Sheets[item.Name];
                    sheet.Visible = XlSheetVisibility.xlSheetVisible;
                    item.IsHidden = false;
                }
            };
            var hideItem = new MenuItem { Header = "隐藏" };
            hideItem.Click += (_, _) =>
            {
                int visibleSheetsCount = 0;
                foreach (Worksheet sheet in ExcelApp.ActiveWorkbook.Sheets)
                {
                    if (sheet.Visible == XlSheetVisibility.xlSheetVisible) visibleSheetsCount++;
                }

                if (ListBoxSheet.SelectedItems.Count >= visibleSheetsCount)
                {
                    MessageBox.Show("无法隐藏全部表格，至少需要显示【1】表格");
                    return;
                }

                foreach (SelfComSheetCollect item in ListBoxSheet.SelectedItems)
                {
                    var sheet = ExcelApp.ActiveWorkbook.Sheets[item.Name];
                    sheet.Visible = XlSheetVisibility.xlSheetHidden;
                    item.IsHidden = true;
                }
            };
            contextMenu.Items.Add(showItem);
            contextMenu.Items.Add(hideItem);

            listBoxItem.ContextMenu = contextMenu;
            listBoxItem.ContextMenu.IsOpen = true;

            e.Handled = true;
        }

        private void ListBoxSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var listBox = (ListBox)sender;
            if (listBox.SelectedItem != null)
            {
                var selectedSheetName = ((SelfComSheetCollect)listBox.SelectedItem).Name;
                var selectedSheet = ExcelApp.ActiveWorkbook.Sheets[selectedSheetName];
                selectedSheet.Activate();
                // 更新 StatusBar
                UpdateStatusBar((SelfComSheetCollect)listBox.SelectedItem);
            }
        }

        public void SetListBoxItemStyle(ListBox listBox)
        {
            Style itemContainerStyle = new Style(typeof(ListBoxItem));

            DataTrigger trigger = new DataTrigger()
            {
                Binding = new System.Windows.Data.Binding("IsHidden"),
                Value = true
            };

            trigger.Setters.Add(new Setter(FontStyleProperty, FontStyles.Italic));
            trigger.Setters.Add(new Setter(ForegroundProperty, System.Windows.Media.Brushes.PapayaWhip));

            itemContainerStyle.Triggers.Add(trigger);

            // 添加 ToolTip
            itemContainerStyle.Setters.Add(new Setter(ToolTipProperty, new System.Windows.Data.Binding("DetailInfo")));

            listBox.ItemContainerStyle = itemContainerStyle;
        }

        // 更新 StatusBar 的方法
        private void UpdateStatusBar(SelfComSheetCollect item)
        {
            StatusBar.Items.Clear();
            var statusBarItem = new StatusBarItem
            {
                Content = 
                    "区域：" + item.UsedRangeSize.Item1 + "行 ," + item.UsedRangeSize.Item2 + "列"
            };
            statusBarItem.ToolTip = statusBarItem.Content; // 设置 ToolTip
            StatusBar.Items.Add(statusBarItem);
        }

    }

}
