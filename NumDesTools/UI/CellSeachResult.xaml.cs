using System.Collections.ObjectModel;
using System.Windows.Controls;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// CellSeachResult.xaml 的交互逻辑
    /// </summary>
    // ReSharper disable once RedundantExtendsListEntry
    public partial class CellSeachResult : UserControl
    {
        public ObservableCollection<SelfCellData> CellDataList { get; set; }

        public CellSeachResult(List<(string, int, int)> list)
        {
            InitializeComponent();
            this.DataContext = this;
            CellDataList = new ObservableCollection<SelfCellData>(list.Select(t => new SelfCellData(t)));
            ListBoxCellData.ItemsSource = CellDataList;
        }

        private void ListBoxCellData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListBoxCellData.SelectedItem is SelfCellData cellData)
            {
                var sheet = NumDesAddIn.App.ActiveSheet;
                var cell = sheet.Cells[cellData.Row, cellData.Column];
                cell.Select();
            }
        }
    }
}
