using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Documents;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// CellSeachResult.xaml 的交互逻辑
    /// </summary>
    // ReSharper disable once RedundantExtendsListEntry
    public partial class SheetCellSeachResult : UserControl
    {
        public ObservableCollection<SelfSheetCellData> CellDataList { get; set; }

        public SheetCellSeachResult(List<(string, int, int, string, string)> list)
        {
            InitializeComponent();
            this.DataContext = this;
            CellDataList = new ObservableCollection<SelfSheetCellData>(list.Select(t => new SelfSheetCellData(t)));
            ListBoxCellData.ItemsSource = CellDataList;
        }

        private void ListBoxCellData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListBoxCellData.SelectedItem is SelfSheetCellData cellData)
            {
                var sheetName = cellData.SheetName;
                var sheet = NumDesAddIn.App.Worksheets[sheetName];

                // 关闭所有打开的备注编辑框
                NumDesAddIn.App.DisplayCommentIndicator = XlCommentDisplayMode.xlNoIndicator;

                sheet.Select();
                var cell = sheet.Cells[cellData.Row, cellData.Column];
                cell.Select();
            }
        }
    }
}

