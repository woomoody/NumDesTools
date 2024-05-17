using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ListBox = System.Windows.Controls.ListBox;
using MenuItem = System.Windows.Controls.MenuItem;
using MessageBox = System.Windows.MessageBox;
using Style = System.Windows.Style;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// SheetListControl.xaml 的交互逻辑
    /// </summary>
    public partial class SheetListControl : UserControl
    {
        public static Application excelApp = NumDesAddIn.App;
        public ObservableCollection<WorksheetWrapper> Sheets { get; } = new ObservableCollection<WorksheetWrapper>();
        public SheetListControl()
        {
            InitializeComponent();
            var worksheets = excelApp.ActiveWorkbook.Sheets.Cast<Worksheet>()
                .Select(x => new WorksheetWrapper { Name = x.Name, IsHidden = x.Visible == XlSheetVisibility.xlSheetHidden });
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
            showItem.Click += (s, args) =>
            {
                foreach (WorksheetWrapper item in ListBoxSheet.SelectedItems)
                {
                    var sheet = excelApp.ActiveWorkbook.Sheets[item.Name];
                    sheet.Visible = XlSheetVisibility.xlSheetVisible;
                    item.IsHidden = false;
                }
            };
            var hideItem = new MenuItem { Header = "隐藏" };
            hideItem.Click += (s, args) =>
            {
                int visibleSheetsCount = 0;
                foreach (Worksheet sheet in excelApp.ActiveWorkbook.Sheets)
                {
                    if (sheet.Visible == XlSheetVisibility.xlSheetVisible) visibleSheetsCount++;
                }

                if (ListBoxSheet.SelectedItems.Count >= visibleSheetsCount)
                {
                    MessageBox.Show("无法隐藏全部表格，至少需要显示【1】表格");
                    return;
                }

                foreach (WorksheetWrapper item in ListBoxSheet.SelectedItems)
                {
                    var sheet = excelApp.ActiveWorkbook.Sheets[item.Name];
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
                var selectedSheetName = ((WorksheetWrapper)listBox.SelectedItem).Name;
                var selectedSheet = excelApp.ActiveWorkbook.Sheets[selectedSheetName];
                selectedSheet.Activate();
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

            listBox.ItemContainerStyle = itemContainerStyle;
        }
    }

}
