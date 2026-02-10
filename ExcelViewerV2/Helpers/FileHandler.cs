using Microsoft.Win32;
using System.Windows;
using ExcelViewerV2.Excel;
using ExcelViewerV2.Controls;

namespace ExcelViewerV2.Helpers
{
    public class FileHandler
    {
        private readonly ExcelLoader _loader;
        private readonly ExcelLikeGrid _grid;
        public bool HasUnsavedChanges { get; private set; }

        public FileHandler(ExcelLoader loader, ExcelLikeGrid grid)
        {
            _loader = loader;
            _grid = grid;
        }

        public void New()
        {
            _loader.CreateNewWorkbook();
            HasUnsavedChanges = false;
            _grid.SetCurrentExcelPath(null);
        }

        public void Open()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Pliki Excel (*.xlsx)|*.xlsx;*.xlsm",
                Title = "Wybierz plik Excel"
            };

            if (dlg.ShowDialog() != true) return;

            _loader.LoadWorkbook(dlg.FileName);
            _loader.LoadAllSheets();

            HasUnsavedChanges = false;
            _grid.SetCurrentExcelPath(dlg.FileName);
        }

        public void Save()
        {
            if (_loader.Save())
            {
                HasUnsavedChanges = false;
                return;
            }

            SaveAs();
        }

        public void SaveAs()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Pliki Excel (*.xlsx)|*.xlsx",
                Title = "Zapisz jako"
            };

            if (dlg.ShowDialog() != true) return;

            if (_loader.SaveAs(dlg.FileName))
            {
                HasUnsavedChanges = false;
                _grid.SetCurrentExcelPath(dlg.FileName);
            }
        }
    }
}
