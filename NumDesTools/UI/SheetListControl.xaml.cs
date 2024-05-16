using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Control = System.Windows.Controls.Control;
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
        //右键显示隐藏
        private void ListBoxSheet_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;
            while (source != null && !(source is ListBoxItem))
            {
                source = VisualTreeHelper.GetParent(source);
            }

            var listBoxItem = source as ListBoxItem;
            if (listBoxItem == null)
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
                    item.IsHidden = false;  // 更新WorksheetWrapper的IsHidden属性
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
                    // 创建一个新的工作表
                    var newSheet = excelApp.ActiveWorkbook.Sheets.Add(After: excelApp.ActiveWorkbook.Sheets[excelApp.ActiveWorkbook.Sheets.Count]);
                    // 创建一个新的WorksheetWrapper并添加到Sheets集合中
                    var newSheetWrapper = new WorksheetWrapper { Name = newSheet.Name, IsHidden = false };
                    Sheets.Add(newSheetWrapper);  // 注意这里改为向Sheets添加新的WorksheetWrapper
                }

                foreach (WorksheetWrapper item in ListBoxSheet.SelectedItems)
                {
                    var sheet = excelApp.ActiveWorkbook.Sheets[item.Name];
                    sheet.Visible = XlSheetVisibility.xlSheetHidden;
                    item.IsHidden = true;  // 更新WorksheetWrapper的IsHidden属性
                }
            };
            contextMenu.Items.Add(showItem);
            contextMenu.Items.Add(hideItem);

            listBoxItem.ContextMenu = contextMenu;
            listBoxItem.ContextMenu.IsOpen = true;

            e.Handled = true;
        }


        //点击Sheet切换显示
        private void ListBoxSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var listBox = (ListBox)sender;
            var selectedSheetName = ((WorksheetWrapper)listBox.SelectedItem).Name;
            var selectedSheet = excelApp.ActiveWorkbook.Sheets[selectedSheetName];
            selectedSheet.Activate();
        }
        //变更listbox样式
        public void SetListBoxItemStyle(ListBox listBox)
        {
            // 添加样式
            Style itemContainerStyle = new Style(typeof(ListBoxItem));

            DataTrigger trigger = new DataTrigger()
            {
                Binding = new System.Windows.Data.Binding("IsHidden"),
                Value = true
            };

            trigger.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Bold));
            trigger.Setters.Add(new Setter(Control.ForegroundProperty, System.Windows.Media.Brushes.Red));

            itemContainerStyle.Triggers.Add(trigger);

            listBox.ItemContainerStyle = itemContainerStyle;
        }


    }
    //自定义表格类
    public class WorksheetWrapper : INotifyPropertyChanged
    {
        private string _name;
        private bool _isHidden;

        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        public bool IsHidden
        {
            get { return _isHidden; }
            set
            {
                if (_isHidden != value)
                {
                    _isHidden = value;
                    OnPropertyChanged(nameof(IsHidden));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
