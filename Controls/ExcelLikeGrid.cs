using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelViewerV2.Model;
using ExcelViewerV2.Editing;
using ExcelViewerV2.Rendering;
using ExcelViewerV2.Core.UndoRedo;

namespace ExcelViewerV2.Controls
{
    public sealed class ExcelLikeGrid : FrameworkElement
    {
        private SpreadsheetModel _model;
        private readonly CellEditorController _editorController;
        private readonly UndoManager _undoManager;
        private CellTextRenderer _textRenderer;

        private HashSet<(int row, int col)> _searchHighlights = new();
        private const int HyperlinkColumn = 6;
        private string _currentExcelPath = "";

        public SpreadsheetModel Model
        {
            get => _model;
            set
            {
                if (value == null) return;

                _model = value;
                _searchHighlights.Clear();

                _editorController.AttachModel(_model);
                _textRenderer = new CellTextRenderer(_model, this);

                InvalidateMeasure();
                InvalidateVisual();
            }
        }

        public ExcelLikeGrid()
        {
            Focusable = true;

            _model = new SpreadsheetModel(100, 26);

            _undoManager = new UndoManager(); // ✅ GLOBALNY UNDO
            _editorController = new CellEditorController(this, _undoManager);
            _editorController.AttachModel(_model);

            AddVisualChild(_editorController.EditorVisual);
            AddLogicalChild(_editorController.EditorVisual);

            _textRenderer = new CellTextRenderer(_model, this);
        }

        // =========================================================
        // 🔥 GLOBALNE UNDO / REDO – EXCEL STYLE
        // =========================================================
        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Z && Keyboard.Modifiers == ModifierKeys.Control)
            {
                // najpierw commit aktualnej edycji
                _editorController.EndEdit();

                _undoManager.Undo();

                InvalidateVisual();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Y && Keyboard.Modifiers == ModifierKeys.Control)
            {
                _editorController.EndEdit();

                _undoManager.Redo();

                InvalidateVisual();
                e.Handled = true;
                return;
            }

            base.OnPreviewKeyDown(e);
        }
        // =========================================================

        public void SetCurrentExcelPath(string excelPath)
        {
            _currentExcelPath = excelPath ?? "";
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            if (_model == null)
                return availableSize;

            double width = 0;
            double height = 0;

            for (int c = 0; c < _model.Columns; c++)
                width += _model.ColumnWidths[c];

            for (int r = 0; r < _model.Rows; r++)
                height += _model.RowHeights[r];

            _editorController.EditorVisual.Measure(availableSize);
            return new Size(width, height);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            _editorController.ArrangeEditor();
            return finalSize;
        }

        protected override void OnRender(DrawingContext dc)
        {
            if (_model == null)
                return;

            double y = 0;

            for (int r = 0; r < _model.Rows; r++)
            {
                double x = 0;
                double rh = _model.RowHeights[r];

                for (int c = 0; c < _model.Columns; c++)
                {
                    double cw = _model.ColumnWidths[c];
                    var cell = _model[r, c];

                    Rect rect = new Rect(x, y, cw, rh);

                    dc.DrawRectangle(cell.Background ?? Brushes.White, null, rect);

                    if (_searchHighlights.Contains((r, c)))
                    {
                        dc.DrawRectangle(
                            new SolidColorBrush(Color.FromArgb(80, 255, 255, 0)),
                            null,
                            rect);
                    }

                    if (_model.ShowGridLines)
                    {
                        dc.DrawRectangle(null, new Pen(Brushes.Gray, 0.5), rect);
                    }

                    x += cw;
                }

                y += rh;
            }

            _textRenderer.Render(dc);
        }

        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            if (_model == null)
                return;

            var hit = HitTestCell(e.GetPosition(this));

            if (hit.row >= 0 && hit.col == HyperlinkColumn)
            {
                var cell = _model[hit.row, HyperlinkColumn];
                if (cell.Hyperlink != null)
                {
                    _editorController.EndEdit();
                    OpenHyperlink(cell.Hyperlink);
                    e.Handled = true;
                    return;
                }
            }

            _editorController.EndEdit();

            if (hit.row >= 0 && hit.col >= 0)
            {
                _editorController.BeginEdit(hit.row, hit.col, hit.rect);
            }

            InvalidateVisual();
            e.Handled = true;
        }

        public void ScrollToCell(int row, int col)
        {
            if (_model == null || row < 0 || col < 0)
                return;

            double y = 0;
            for (int r = 0; r < row; r++)
                y += _model.RowHeights[r];

            double x = 0;
            for (int c = 0; c < col; c++)
                x += _model.ColumnWidths[c];

            ScrollViewer sv = FindScrollViewer();
            if (sv != null)
            {
                sv.ScrollToVerticalOffset(y);
                sv.ScrollToHorizontalOffset(x);
            }

            Rect rect = new Rect(
                x,
                y,
                _model.ColumnWidths[col],
                _model.RowHeights[row]);

            _editorController.BeginEdit(row, col, rect);
        }

        private ScrollViewer FindScrollViewer()
        {
            DependencyObject parent = this;

            while (parent != null)
            {
                if (parent is ScrollViewer sv)
                    return sv;

                parent = VisualTreeHelper.GetParent(parent);
            }

            return null;
        }

        private void OpenHyperlink(Uri uri)
        {
            try
            {
                if (!File.Exists(uri.LocalPath))
                {
                    MessageBox.Show($"Plik nie istnieje:\n{uri.LocalPath}");
                    return;
                }

                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = uri.LocalPath,
                        UseShellExecute = true
                    });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Błąd przy otwieraniu linku:\n{ex.Message}");
            }
        }

        internal (int row, int col, Rect rect) HitTestCell(Point p)
        {
            double y = 0;

            for (int r = 0; r < _model.Rows; r++)
            {
                double x = 0;

                for (int c = 0; c < _model.Columns; c++)
                {
                    Rect rect = new Rect(
                        x,
                        y,
                        _model.ColumnWidths[c],
                        _model.RowHeights[r]);

                    if (rect.Contains(p))
                        return (r, c, rect);

                    x += rect.Width;
                }

                y += _model.RowHeights[r];
            }

            return (-1, -1, Rect.Empty);
        }

        internal void EndEdit()
        {
            _editorController.EndEdit();
        }

        protected override int VisualChildrenCount => 1;

        protected override Visual GetVisualChild(int index)
        {
            if (index == 0)
                return _editorController.EditorVisual;

            throw new ArgumentOutOfRangeException();
        }

        public void Undo()
{
    _editorController.EndEdit();
    _undoManager.Undo();
    InvalidateVisual();
}

public void Redo()
{
    _editorController.EndEdit();
    _undoManager.Redo();
    InvalidateVisual();
}
    }
}
