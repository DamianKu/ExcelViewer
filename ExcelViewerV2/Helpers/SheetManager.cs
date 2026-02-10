using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows;
using Microsoft.Win32;
using ExcelViewerV2.Excel;
using ExcelViewerV2.Controls;

namespace ExcelViewerV2.Helpers
{
    public class SheetManager
    {
        private readonly ExcelLoader _loader;
        private readonly Panel _sheetPanel;
        private readonly ScrollViewer _scrollViewer;
        private readonly Button _leftArrow;
        private readonly Button _rightArrow;

        private int _scrollIndex = 0;
        private const double MinSheetButtonWidth = 60;
        private const double VisibleWidthRatio = 0.90;

        public string ActiveSheetName { get; set; }

        // Event powiadamiający MainWindow o zmianie aktywnego arkusza
        public event Action<string> ActiveSheetChanged;

        public SheetManager(ExcelLoader loader, Panel sheetPanel, ScrollViewer scrollViewer, Button leftArrow, Button rightArrow)
        {
            _loader = loader;
            _sheetPanel = sheetPanel;
            _scrollViewer = scrollViewer;
            _leftArrow = leftArrow;
            _rightArrow = rightArrow;
        }

        public void RefreshSheetButtons()
        {
            if (_sheetPanel == null || _scrollViewer == null) return;

            _sheetPanel.Children.Clear();
            var visibleSheets = _loader.GetAllSheetModels().Values
                .Where(s => s.IsVisible)
                .Select(s => s.Name)
                .ToList();

            int totalSheets = visibleSheets.Count;
            int visibleCount = GetVisibleSheetsCount();

            _scrollIndex = Math.Max(0, Math.Min(_scrollIndex, totalSheets - visibleCount));
            double availableWidth = _scrollViewer.ActualWidth * VisibleWidthRatio;
            double buttonWidth = availableWidth / visibleCount;

            for (int i = _scrollIndex; i < totalSheets && i < _scrollIndex + visibleCount; i++)
            {
                string sheetName = visibleSheets[i];
                var btn = new Button
                {
                    Content = sheetName,
                    Width = buttonWidth,
                    Margin = new Thickness(1, 0, 1, 0),
                    FontWeight = sheetName == ActiveSheetName ? FontWeights.Bold : FontWeights.Normal
                };

                btn.Click += (_, __) =>
                {
                    ActiveSheetName = sheetName;
                    ActiveSheetChanged?.Invoke(sheetName); // <-- powiadom MainWindow
                    EnsureActiveSheetButtonVisible();
                };

                btn.MouseRightButtonUp += SheetButton_MouseRightButtonUp;
                _sheetPanel.Children.Add(btn);
            }

            _leftArrow.IsEnabled = _scrollIndex > 0;
            _rightArrow.IsEnabled = _scrollIndex + visibleCount < totalSheets;
        }

        private int GetVisibleSheetsCount()
        {
            if (_scrollViewer == null || _scrollViewer.ActualWidth <= 0) return 1;
            double availableWidth = _scrollViewer.ActualWidth * VisibleWidthRatio;
            return Math.Max(1, (int)(availableWidth / MinSheetButtonWidth));
        }

        public void EnsureActiveSheetButtonVisible()
        {
            var visible = _loader.GetAllSheetModels().Values
                .Where(s => s.IsVisible)
                .Select((s, i) => new { s.Name, i })
                .ToList();

            int index = visible.FirstOrDefault(x => x.Name == ActiveSheetName)?.i ?? 0;
            int visibleCount = GetVisibleSheetsCount();

            if (index < _scrollIndex)
                _scrollIndex = index;
            else if (index >= _scrollIndex + visibleCount)
                _scrollIndex = index - visibleCount + 1;

            RefreshSheetButtons();
        }

        private void SheetButton_MouseRightButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
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

        public void ScrollLeft()
        {
            if (_scrollIndex > 0)
            {
                _scrollIndex--;
                RefreshSheetButtons();
            }
        }

        public void ScrollRight()
        {
            int total = _loader.GetAllSheetModels().Values.Count(s => s.IsVisible);
            int visible = GetVisibleSheetsCount();

            if (_scrollIndex + visible < total)
            {
                _scrollIndex++;
                RefreshSheetButtons();
            }
        }
    }
}
