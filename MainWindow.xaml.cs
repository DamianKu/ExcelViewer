using Microsoft.Win32;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelViewerV2.Excel;
using ExcelViewerV2.Controls;
using ExcelViewerV2.Model;
using ExcelViewerV2.Printing;
using ExcelViewerV2.Search;

namespace ExcelViewerV2
{
    public partial class MainWindow : Window
    {
        private readonly ExcelLoader _loader = new ExcelLoader();

        private int _scrollIndex = 0;
        private string _activeSheetName;
        private FindDialog _findDialog;

        private bool _hasUnsavedChanges;

        private const double MinSheetButtonWidth = 60;
        private const double VisibleWidthRatio = 0.90;

        public MainWindow()
        {
            InitializeComponent();

            Loaded += (_, __) =>
            {
                RefreshSheetButtons();
                UpdateHorizontalScrollBar();
            };

            SizeChanged += (_, __) =>
            {
                RefreshSheetButtons();
                UpdateHorizontalScrollBar();
            };
        }

        // ======================
        // SKRÓTY
        // ======================
        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.F && Keyboard.Modifiers == ModifierKeys.Control)
            {
                ShowFindDialog();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
            {
                SaveButton_Click(this, new RoutedEventArgs());
                e.Handled = true;
                return;
            }

            if (e.Key == Key.S && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                SaveAsButton_Click(this, new RoutedEventArgs());
                e.Handled = true;
                return;
            }

            base.OnPreviewKeyDown(e);
        }

        // ======================
        // WYSZUKIWANIE
        // ======================
        private void ShowFindDialog()
        {
            if (ExcelGrid?.Model == null) return;

            if (_findDialog == null || !_findDialog.IsVisible)
            {
                _findDialog = new FindDialog(
                    this,
                    _loader.GetAllSheetModels().Values
                        .Where(s => s.IsVisible)
                        .ToList()
                )
                { Owner = this };

                _findDialog.Show();
            }
            else
            {
                _findDialog.Activate();
            }
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e) => ShowFindDialog();

        public void GoToSearchResult(SearchResult result)
        {
            if (result == null) return;

            var sheet = _loader.GetSheetModel(result.SheetName);
            if (sheet == null || !sheet.IsVisible) return;

            _activeSheetName = sheet.Name;

            DisplayActiveSheet();
            RefreshSheetButtons();
            EnsureActiveSheetButtonVisible();

            ExcelGrid.ScrollToCell(result.Row, result.Column);
        }

        // ======================
        // NOWY
        // ======================
        private void NewButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscardChanges()) return;

            _loader.CreateNewWorkbook();

            _scrollIndex = 0;
            _activeSheetName = _loader.GetAllSheetModels().Values.First().Name;
            _hasUnsavedChanges = false;

            ExcelGrid.SetCurrentExcelPath(null);

            DisplayActiveSheet();
            RefreshSheetButtons();
        }

        // ======================
        // OTWÓRZ
        // ======================
        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscardChanges()) return;

            var dlg = new OpenFileDialog
            {
                Filter = "Pliki Excel (*.xlsx)|*.xlsx;*.xlsm",
                Title = "Wybierz plik Excel"
            };

            if (dlg.ShowDialog() != true) return;

            _loader.LoadWorkbook(dlg.FileName);
            _loader.LoadAllSheets();

            _scrollIndex = 0;
            _activeSheetName = _loader.GetAllSheetModels().Values
                .FirstOrDefault(s => s.IsVisible)?.Name;

            _hasUnsavedChanges = false;

            ExcelGrid.SetCurrentExcelPath(dlg.FileName);

            DisplayActiveSheet();
            RefreshSheetButtons();
            UpdateHorizontalScrollBar();
            RefreshColumnHeaders();
            RefreshRowHeaders();
        }

        // ======================
        // ZAPIS
        // ======================
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (_loader.Save())
            {
                _hasUnsavedChanges = false;
                return;
            }

            SaveAsButton_Click(sender, e);
        }

        private void SaveAsButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Pliki Excel (*.xlsx)|*.xlsx",
                Title = "Zapisz jako"
            };

            if (dlg.ShowDialog() != true) return;

            if (_loader.SaveAs(dlg.FileName))
            {
                _hasUnsavedChanges = false;
                ExcelGrid.SetCurrentExcelPath(dlg.FileName);
            }
        }

        private bool ConfirmDiscardChanges()
        {
            if (!_hasUnsavedChanges) return true;

            var res = MessageBox.Show(
                "Masz niezapisane zmiany. Czy chcesz je zapisać?",
                "Niezapisane zmiany",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Warning);

            if (res == MessageBoxResult.Yes)
            {
                SaveButton_Click(this, new RoutedEventArgs());
                return !_hasUnsavedChanges;
            }

            return res == MessageBoxResult.No;
        }

        // ======================
        // NAGŁÓWKI
        // ======================
        private void RefreshColumnHeaders()
        {
            if (ExcelGrid?.Model == null) return;
            ColumnHeaderPanel.Children.Clear();

            for (int c = 0; c < ExcelGrid.Model.Columns; c++)
            {
                ColumnHeaderPanel.Children.Add(new TextBlock
                {
                    Text = GetExcelColumnName(c),
                    Width = ExcelGrid.Model.ColumnWidths[c],
                    TextAlignment = TextAlignment.Center,
                    FontWeight = FontWeights.Bold,
                    Background = Brushes.LightGray,
                    Margin = new Thickness(0.5),
                    Padding = new Thickness(2, 0, 2, 0)
                });
            }
        }

        private void RefreshRowHeaders()
        {
            if (ExcelGrid?.Model == null) return;
            RowHeaderPanel.Children.Clear();

            for (int r = 0; r < ExcelGrid.Model.Rows; r++)
            {
                RowHeaderPanel.Children.Add(new TextBlock
                {
                    Text = (r + 1).ToString(),
                    Height = ExcelGrid.Model.RowHeights[r],
                    Width = 30,
                    TextAlignment = TextAlignment.Center,
                    FontWeight = FontWeights.Bold,
                    Background = Brushes.LightGray,
                    Margin = new Thickness(0.5),
                    Padding = new Thickness(2, 0, 2, 0)
                });
            }
        }

        private static string GetExcelColumnName(int index)
        {
            string name = "";
            do
            {
                name = (char)('A' + (index % 26)) + name;
                index = (index / 26) - 1;
            } while (index >= 0);
            return name;
        }

        // ======================
        // ARKUSZE
        // ======================
        private int GetVisibleSheetsCount()
        {
            if (SheetScrollViewer == null || SheetScrollViewer.ActualWidth <= 0) return 1;
            double availableWidth = SheetScrollViewer.ActualWidth * VisibleWidthRatio;
            return Math.Max(1, (int)(availableWidth / MinSheetButtonWidth));
        }

        private void RefreshSheetButtons()
        {
            if (SheetPanel == null || SheetScrollViewer == null) return;

            SheetPanel.Children.Clear();

            var visibleSheets = _loader.GetAllSheetModels().Values
                .Where(s => s.IsVisible)
                .Select(s => s.Name)
                .ToList();

            int totalSheets = visibleSheets.Count;
            int visibleCount = GetVisibleSheetsCount();

            _scrollIndex = Math.Max(0, Math.Min(_scrollIndex, totalSheets - visibleCount));

            double availableWidth = SheetScrollViewer.ActualWidth * VisibleWidthRatio;
            double buttonWidth = availableWidth / visibleCount;

            for (int i = _scrollIndex; i < totalSheets && i < _scrollIndex + visibleCount; i++)
            {
                string sheetName = visibleSheets[i];
                var btn = new Button
                {
                    Content = sheetName,
                    Width = buttonWidth,
                    Margin = new Thickness(1, 0, 1, 0),
                    FontWeight = sheetName == _activeSheetName
                        ? FontWeights.Bold
                        : FontWeights.Normal
                };
                btn.Click += SheetButton_Click;
                btn.MouseRightButtonUp += SheetButton_MouseRightButtonUp;
                SheetPanel.Children.Add(btn);
            }

            LeftArrow.IsEnabled = _scrollIndex > 0;
            RightArrow.IsEnabled = _scrollIndex + visibleCount < totalSheets;
        }

        private void SheetButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn)
            {
                _activeSheetName = btn.Content?.ToString();
                DisplayActiveSheet();
                RefreshSheetButtons();
                EnsureActiveSheetButtonVisible();
            }
        }

        private void SheetButton_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Button btn) return;
            string sheetName = btn.Content?.ToString();
            if (sheetName == null) return;

            var sheet = _loader.GetSheetModel(sheetName);
            if (sheet == null) return;

            var menu = new ContextMenu();

            if (sheet.IsVisible)
            {
                var hide = new MenuItem { Header = "Ukryj arkusz" };
                hide.Click += (_, __) =>
                {
                    sheet.IsVisible = false;
                    RefreshSheetButtons();
                };
                menu.Items.Add(hide);
            }

            var hidden = _loader.GetAllSheetModels().Values.Where(s => !s.IsVisible).ToList();
            if (hidden.Count > 0)
            {
                menu.Items.Add(new Separator());
                var show = new MenuItem { Header = "Odkryj arkusz" };
                menu.Items.Add(show);

                foreach (var h in hidden)
                {
                    var item = new MenuItem { Header = h.Name };
                    item.Click += (_, __) =>
                    {
                        h.IsVisible = true;
                        RefreshSheetButtons();
                    };
                    show.Items.Add(item);
                }
            }

            menu.IsOpen = true;
        }

        private void EnsureActiveSheetButtonVisible()
        {
            var visible = _loader.GetAllSheetModels().Values
                .Where(s => s.IsVisible)
                .Select((s, i) => new { s.Name, i })
                .ToList();

            int index = visible.FirstOrDefault(x => x.Name == _activeSheetName)?.i ?? 0;
            int visibleCount = GetVisibleSheetsCount();

            if (index < _scrollIndex)
                _scrollIndex = index;
            else if (index >= _scrollIndex + visibleCount)
                _scrollIndex = index - visibleCount + 1;

            RefreshSheetButtons();
        }

        private void DisplayActiveSheet()
        {
            if (_activeSheetName == null) return;

            var model = _loader.GetSheetModel(_activeSheetName);
            if (model == null) return;

            ExcelGrid.Model = model;
            RefreshColumnHeaders();
            RefreshRowHeaders();
        }

        // ======================
        // SCROLL
        // ======================
        private void LeftArrow_Click(object sender, RoutedEventArgs e)
        {
            if (_scrollIndex > 0)
            {
                _scrollIndex--;
                RefreshSheetButtons();
            }
        }

        private void RightArrow_Click(object sender, RoutedEventArgs e)
        {
            int total = _loader.GetAllSheetModels().Values.Count(s => s.IsVisible);
            int visible = GetVisibleSheetsCount();

            if (_scrollIndex + visible < total)
            {
                _scrollIndex++;
                RefreshSheetButtons();
            }
        }

        private void ExcelHorizontalScrollBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            ScrollViewerGrid?.ScrollToHorizontalOffset(e.NewValue);
        }

        private void UpdateHorizontalScrollBar()
        {
            if (ScrollViewerGrid == null || ExcelHorizontalScrollBar == null) return;

            ExcelHorizontalScrollBar.Minimum = 0;
            ExcelHorizontalScrollBar.Maximum =
                Math.Max(0, ScrollViewerGrid.ExtentWidth - ScrollViewerGrid.ViewportWidth);
            ExcelHorizontalScrollBar.ViewportSize = ScrollViewerGrid.ViewportWidth;
            ExcelHorizontalScrollBar.SmallChange = 10;
            ExcelHorizontalScrollBar.LargeChange = ExcelHorizontalScrollBar.ViewportSize;
        }

        private void ScrollViewerGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            ColumnHeaderScrollViewer.ScrollToHorizontalOffset(e.HorizontalOffset);
            RowHeaderScrollViewer.ScrollToVerticalOffset(e.VerticalOffset);
        }

        private void Undo_Click(object sender, RoutedEventArgs e)
{
    ExcelGrid?.Undo();
}

private void Redo_Click(object sender, RoutedEventArgs e)
{
    ExcelGrid?.Redo();
}

        // ======================
        // DRUK
        // ======================
        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            if (ExcelGrid.Model == null)
            {
                MessageBox.Show("Brak aktywnego arkusza do wydruku.",
                    "Drukowanie",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            new SpreadsheetPrinter().Print(ExcelGrid.Model);
        }
        
    }
}
