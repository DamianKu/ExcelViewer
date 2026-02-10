using System.Collections.Generic;
using System.Windows;
using ExcelViewerV2.Controls;
using ExcelViewerV2.Model;
using ExcelViewerV2.Helpers;
using ExcelViewerV2.Search;

namespace ExcelViewerV2.Services
{
    public class DialogService : IDialogService
    {
        public void ShowFindDialog(List<SpreadsheetModel> sheets, SheetManager sheetManager, ExcelLikeGrid grid)
        {
            var dlg = new FindDialog((MainWindow)Application.Current.MainWindow, sheets);
            dlg.Show();
        }

        public void ShowInfo(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public bool Confirm(string message, string title)
        {
            var res = MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
            return res == MessageBoxResult.Yes;
        }
    }
}
