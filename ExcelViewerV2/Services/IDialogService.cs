using System.Collections.Generic;
using ExcelViewerV2.Controls;
using ExcelViewerV2.Model;
using ExcelViewerV2.Helpers;


namespace ExcelViewerV2.Services
{
    public interface IDialogService
    {
        void ShowFindDialog(List<SpreadsheetModel> sheets, SheetManager sheetManager, ExcelLikeGrid grid);
        void ShowInfo(string message, string title);
        bool Confirm(string message, string title);
    }
}
