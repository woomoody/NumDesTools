using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// CellSeachResult.xaml 的交互逻辑
    /// </summary>
    // ReSharper disable once RedundantExtendsListEntry
    public partial class CellSeachResult : UserControl
    {
        public ObservableCollection<CellData> CellDataList { get; set; }

        public CellSeachResult(List<(string, int, int)> list)
        {
            InitializeComponent();
            this.DataContext = this;
            CellDataList = new ObservableCollection<CellData>(list.Select(t => new CellData(t)));
            ListBoxCellData.ItemsSource = CellDataList;
        }

        private void ListBoxCellData_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListBoxCellData.SelectedItem is CellData cellData)
            {
                var sheet = NumDesAddIn.App.ActiveSheet;
                var cell = sheet.Cells[cellData.Row, cellData.Column];
                cell.Select();
            }
        }
    }
}
