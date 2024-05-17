using System.Windows;
using System.Windows.Controls;
using MenuItem = System.Windows.Controls.MenuItem;
using Window = System.Windows.Window;

namespace NumDesTools.UI
{
    /// <summary>
    /// TestWPF.xaml 的交互逻辑
    /// </summary>
    public partial class SheetListWindow : Window
    {
        public SheetListWindow()
        {
            InitializeComponent();

            var excelApp = NumDesAddIn.App;
            var sheets = excelApp.ActiveWorkbook.Sheets.Cast<Worksheet>().ToList();
            ListBoxSheet.ItemsSource = sheets;
        }

        private void ShowSheet_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = (MenuItem)sender;
            var contextMenu = (ContextMenu)menuItem.Parent;
            var textBlock = (TextBlock)contextMenu.PlacementTarget;
            var sheet = (Worksheet)textBlock.DataContext;

            sheet.Visible = XlSheetVisibility.xlSheetVisible;
        }

        private void HideSheet_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = (MenuItem)sender;
            var contextMenu = (ContextMenu)menuItem.Parent;
            var textBlock = (TextBlock)contextMenu.PlacementTarget;
            var sheet = (Worksheet)textBlock.DataContext;

            sheet.Visible = XlSheetVisibility.xlSheetHidden;
        }
    }


}
