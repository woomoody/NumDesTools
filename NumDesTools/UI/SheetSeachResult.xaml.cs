using System.Collections.ObjectModel;
using System.Windows.Controls;
using ListBox = System.Windows.Controls.ListBox;

namespace NumDesTools.UI
{
    /// <summary>
    /// SheetSeachResult.xaml 的交互逻辑
    /// </summary>
    public partial class SheetSeachResult
    {
        public ObservableCollection<WorkBookSearchCollect> TargetSheetList { get; set; }

        public SheetSeachResult(List<(string, string, int, string)> list)
        {
            InitializeComponent();
            this.DataContext = this;
            TargetSheetList = new ObservableCollection<WorkBookSearchCollect>(
                list.Select(t => new WorkBookSearchCollect(t))
            );
        }

        private void ListBoxWorkBook_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var listBox = (ListBox)sender;
            if (listBox.SelectedItem != null)
            {
                var selectedWorkBook = (WorkBookSearchCollect)listBox.SelectedItem;
                PubMetToExcel.OpenExcelAndSelectCell(
                    selectedWorkBook.FilePath,
                    selectedWorkBook.SheetName,
                    selectedWorkBook.CellCol + selectedWorkBook.CellRow
                );
            }
        }
    }
}
