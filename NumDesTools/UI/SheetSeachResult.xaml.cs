﻿using System.Collections.ObjectModel;
using System.Windows.Controls;
using ListBox = System.Windows.Controls.ListBox;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// SheetSeachResult.xaml 的交互逻辑
    /// </summary>
    // ReSharper disable once RedundantExtendsListEntry
    public partial class SheetSeachResult:UserControl
    {
        public ObservableCollection<SelfWorkBookSearchCollect> TargetSheetList { get; set; }

        public SheetSeachResult(List<(string, string, int, string)> list)
        {
            InitializeComponent();
            DataContext = this;
            TargetSheetList = new ObservableCollection<SelfWorkBookSearchCollect>(
                list.Select(t => new SelfWorkBookSearchCollect(t))
            );
        }
        //这个事件的问题是打开新的再关闭后，再次打开无法点击ListBox
        //但是如果采用单击listbox事件会导致打开的Excel文件不在前景，暂时没想到更好的办法
        private void ListBoxWorkBook_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var listBox = (ListBox)sender;
            if (listBox.SelectedItem != null)
            {
                var selectedWorkBook = (SelfWorkBookSearchCollect)listBox.SelectedItem;
                // 关闭所有打开的备注编辑框，不隐藏角标
                NumDesAddIn.App.DisplayCommentIndicator = XlCommentDisplayMode.xlCommentIndicatorOnly;

                PubMetToExcel.OpenExcelAndSelectCell(
                    selectedWorkBook.FilePath,
                    selectedWorkBook.SheetName,
                    selectedWorkBook.CellCol + selectedWorkBook.CellRow
                );
                // 手动清空 SelectedItem，支持重复点击
                listBox.SelectedItem = null;
            }
        }
    }
}
