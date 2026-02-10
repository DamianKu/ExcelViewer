using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelViewerV2.Controls;
using ExcelViewerV2.Excel;
using ExcelViewerV2.Helpers;
using ExcelViewerV2.Search;
using ExcelViewerV2.ViewModels;
using ExcelViewerV2.Services;
using ExcelViewerV2.Rendering;   // <<< DODANE

namespace ExcelViewerV2
{
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        private SheetManager _sheetManager;   // <<< DODANE

        public MainWindow()
        {
            InitializeComponent();

            // =======================
            // INICJALIZACJA LOGIKI
            // =======================
            var loader = new ExcelLoader();
            var fileHandler = new FileHandler(loader, ExcelGrid);

            _sheetManager = new SheetManager(
                loader,
                SheetPanel,
                SheetScrollViewer,
                LeftArrow,
                RightArrow);

            var dialogService = new DialogService();

            _viewModel = new MainViewModel(
                dialogService,
                ExcelGrid,
                _sheetManager,
                fileHandler,
                loader);

            DataContext = _viewModel;

            // <<< KLUCZOWE: odśwież nagłówki przy zmianie arkusza
            _sheetManager.ActiveSheetChanged += _ =>
            {
                RefreshHeaders();
            };

            Loaded += (_, __) =>
            {
                _sheetManager.RefreshSheetButtons();
                UpdateHorizontalScrollBar();
                RefreshHeaders();   // <<< DODANE
            };

            SizeChanged += (_, __) =>
            {
                _sheetManager.RefreshSheetButtons();
                UpdateHorizontalScrollBar();
                RefreshHeaders();   // <<< DODANE
            };
        }

// ======================
// REFRESH HEADERÓW
// ======================
private void RefreshHeaders()
{
    if (ExcelGrid?.Model == null)
    {
        ColumnHeaderScrollViewer.Visibility = Visibility.Collapsed;
        RowHeaderScrollViewer.Visibility = Visibility.Collapsed;
        ScrollViewerGrid.Visibility = Visibility.Collapsed;
        TopLeftCorner.Visibility = Visibility.Collapsed; // <<< ukryj
        return;
    }

    ColumnHeaderScrollViewer.Visibility = Visibility.Visible;
    RowHeaderScrollViewer.Visibility = Visibility.Visible;
    ScrollViewerGrid.Visibility = Visibility.Visible;
    TopLeftCorner.Visibility = Visibility.Visible; // <<< pokaż

    HeaderRenderer.RefreshColumnHeaders(ExcelGrid, ColumnHeaderPanel);
    HeaderRenderer.RefreshRowHeaders(ExcelGrid, RowHeaderPanel);
}

        // ======================
        // SKRÓTY KLAWIATUROWE
        // ======================
        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.F && Keyboard.Modifiers == ModifierKeys.Control)
            {
                _viewModel.SearchCommand.Execute(null);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
            {
                _viewModel.SaveCommand.Execute(null);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.S && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                _viewModel.SaveAsCommand.Execute(null);
                e.Handled = true;
                return;
            }

            base.OnPreviewKeyDown(e);
        }

        // ======================
        // SCROLL
        // ======================
        private void ExcelHorizontalScrollBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            ScrollViewerGrid?.ScrollToHorizontalOffset(e.NewValue);
            RefreshHeaders();   // <<< DODANE
        }

        private void UpdateHorizontalScrollBar()
        {
            if (ScrollViewerGrid == null || ExcelHorizontalScrollBar == null)
                return;

            ExcelHorizontalScrollBar.Minimum = 0;
            ExcelHorizontalScrollBar.Maximum =
                System.Math.Max(0,
                ScrollViewerGrid.ExtentWidth - ScrollViewerGrid.ViewportWidth);

            ExcelHorizontalScrollBar.ViewportSize =
                ScrollViewerGrid.ViewportWidth;

            ExcelHorizontalScrollBar.SmallChange = 10;
            ExcelHorizontalScrollBar.LargeChange =
                ExcelHorizontalScrollBar.ViewportSize;
        }

        private void ScrollViewerGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            ColumnHeaderScrollViewer.ScrollToHorizontalOffset(e.HorizontalOffset);
            RowHeaderScrollViewer.ScrollToVerticalOffset(e.VerticalOffset);

            RefreshHeaders();   // <<< DODANE
        }

        // ======================
        // SEARCH RESULT
        // ======================
        public void GoToSearchResult(SearchResult result)
        {
            _viewModel.GoToSearchResult(result);
        }
    }
}
