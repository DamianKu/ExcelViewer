using System.Linq;
using System.Windows.Input;
using ExcelViewerV2.Controls;
using ExcelViewerV2.Excel;
using ExcelViewerV2.Helpers;
using ExcelViewerV2.Printing;
using ExcelViewerV2.Search;
using ExcelViewerV2.Services;

namespace ExcelViewerV2.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private readonly ExcelLoader _loader;
        private readonly FileHandler _fileHandler;
        private readonly SheetManager _sheetManager;
        private readonly ExcelLikeGrid _excelGrid;
        private readonly IDialogService _dialogService;

        #region Commands
        public ICommand NewCommand { get; }
        public ICommand OpenCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand SaveAsCommand { get; }
        public ICommand SearchCommand { get; }
        public ICommand UndoCommand { get; }
        public ICommand RedoCommand { get; }
        public ICommand PrintCommand { get; }
        public ICommand ScrollLeftCommand { get; }
        public ICommand ScrollRightCommand { get; }
        #endregion

        public MainViewModel(
            IDialogService dialogService,
            ExcelLikeGrid excelGrid,
            SheetManager sheetManager,
            FileHandler fileHandler,
            ExcelLoader loader)
        {
            _dialogService = dialogService;
            _excelGrid = excelGrid;
            _sheetManager = sheetManager;
            _fileHandler = fileHandler;
            _loader = loader;

            _sheetManager.ActiveSheetChanged += _ => DisplayActiveSheet();

            NewCommand = new RelayCommand(_ => New());
            OpenCommand = new RelayCommand(_ => Open());
            SaveCommand = new RelayCommand(_ => Save());
            SaveAsCommand = new RelayCommand(_ => SaveAs());
            SearchCommand = new RelayCommand(_ => ShowFindDialog());
            UndoCommand = new RelayCommand(_ => _excelGrid?.Undo());
            RedoCommand = new RelayCommand(_ => _excelGrid?.Redo());
            PrintCommand = new RelayCommand(_ => Print());
            ScrollLeftCommand = new RelayCommand(_ => ScrollLeft());
            ScrollRightCommand = new RelayCommand(_ => ScrollRight());
        }

        private void New()
        {
            if (!ConfirmDiscardChanges()) return;

            _fileHandler.New();
            var firstSheet = _loader.GetAllSheetModels().Values.FirstOrDefault();
            if (firstSheet != null)
                _sheetManager.ActiveSheetName = firstSheet.Name;

            DisplayActiveSheet();
            _sheetManager.RefreshSheetButtons();
        }

        private void Open()
        {
            if (!ConfirmDiscardChanges()) return;

            _fileHandler.Open();
            var firstVisible = _loader.GetAllSheetModels().Values.FirstOrDefault(s => s.IsVisible);
            if (firstVisible != null)
                _sheetManager.ActiveSheetName = firstVisible.Name;

            DisplayActiveSheet();
            _sheetManager.RefreshSheetButtons();
        }

        private void Save() => _fileHandler.Save();
        private void SaveAs() => _fileHandler.SaveAs();

        private void ShowFindDialog()
        {
            if (_excelGrid?.Model == null) return;

            _dialogService.ShowFindDialog(
                _loader.GetAllSheetModels().Values.Where(s => s.IsVisible).ToList(),
                _sheetManager,
                _excelGrid);
        }

        public void GoToSearchResult(SearchResult result)
        {
            if (result == null) return;

            var sheet = _loader.GetSheetModel(result.SheetName);
            if (sheet == null || !sheet.IsVisible) return;

            _sheetManager.ActiveSheetName = sheet.Name;
            DisplayActiveSheet();
            _sheetManager.RefreshSheetButtons();
            _excelGrid.ScrollToCell(result.Row, result.Column);
        }

        private void Print()
        {
            if (_excelGrid.Model == null)
            {
                _dialogService.ShowInfo("Brak aktywnego arkusza do wydruku.", "Drukowanie");
                return;
            }

            new SpreadsheetPrinter().Print(_excelGrid.Model);
        }

        private void DisplayActiveSheet()
        {
            if (_sheetManager.ActiveSheetName == null) return;

            var model = _loader.GetSheetModel(_sheetManager.ActiveSheetName);
            if (model == null) return;

            _excelGrid.Model = model;
        }

        private bool ConfirmDiscardChanges()
        {
            if (!_fileHandler.HasUnsavedChanges) return true;

            return _dialogService.Confirm("Masz niezapisane zmiany. Czy chcesz je zapisać?", "Niezapisane zmiany");
        }

        public void ScrollLeft() => _sheetManager.ScrollLeft();
        public void ScrollRight() => _sheetManager.ScrollRight();
    }
}
