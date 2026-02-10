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

        private double[] _rowOffsets;
        private double[] _colOffsets;

        public double HorizontalOffset { get; private set; }
        public double VerticalOffset { get; private set; }

        public SpreadsheetModel Model
        {
            get => _model;
            set
            {
                _model = value;

                if (_model == null)
                {
                    _searchHighlights.Clear();
                    _textRenderer = null;
                    _rowOffsets = null;
                    _colOffsets = null;
                    InvalidateVisual();
                    return;
                }

                _searchHighlights.Clear();
                _editorController.AttachModel(_model);
                _textRenderer = new CellTextRenderer(_model, this);

                BuildOffsets();
                HookScrollViewer();

                InvalidateMeasure();
                InvalidateVisual();
            }
        }

        public ExcelLikeGrid()
        {
            Focusable = true;

            Loaded += (s, e) => HookScrollViewer();

            _undoManager = new UndoManager();
            _editorController = new CellEditorController(this, _undoManager);

            AddVisualChild(_editorController.EditorVisual);
            AddLogicalChild(_editorController.EditorVisual);

            // nie tworzymy domyślnego modelu
            _model = null;
            _textRenderer = null;
        }

        private void BuildOffsets()
        {
            if (_model == null) return;

            _rowOffsets = new double[_model.Rows + 1];
            _colOffsets = new double[_model.Columns + 1];

            for (int i = 1; i <= _model.Rows; i++)
                _rowOffsets[i] = _rowOffsets[i - 1] + _model.RowHeights[i - 1];

            for (int i = 1; i <= _model.Columns; i++)
                _colOffsets[i] = _colOffsets[i - 1] + _model.ColumnWidths[i - 1];
        }

        private int FindIndex(double[] offsets, double value)
        {
            if (offsets == null) return -1;

            int lo = 0;
            int hi = offsets.Length - 1;

            while (lo < hi)
            {
                int mid = (lo + hi) / 2;
                if (offsets[mid] <= value)
                    lo = mid + 1;
                else
                    hi = mid;
            }

            return Math.Max(0, lo - 1);
        }

        private void HookScrollViewer()
        {
            var sv = FindScrollViewer();
            if (sv == null) return;

            sv.ScrollChanged -= OnScrollChanged;
            sv.ScrollChanged += OnScrollChanged;

            HorizontalOffset = sv.HorizontalOffset;
            VerticalOffset = sv.VerticalOffset;
            InvalidateVisual();
        }

        private void OnScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            HorizontalOffset = e.HorizontalOffset;
            VerticalOffset = e.VerticalOffset;
            InvalidateVisual();
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Z && Keyboard.Modifiers == ModifierKeys.Control)
            {
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

        public void SetCurrentExcelPath(string excelPath)
        {
            _currentExcelPath = excelPath ?? "";
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            if (_model == null || _colOffsets == null || _rowOffsets == null)
                return new Size(0, 0);

            double width = _colOffsets[^1];
            double height = _rowOffsets[^1];

            if (!double.IsInfinity(availableSize.Width))
                width = availableSize.Width;

            if (!double.IsInfinity(availableSize.Height))
                height = availableSize.Height;

            return new Size(width, height);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            _editorController.ArrangeEditor();
            return finalSize;
        }

        protected override void OnRender(DrawingContext dc)
        {
            if (_model == null || _rowOffsets == null || _colOffsets == null)
                return;

            double viewW = RenderSize.Width;
            double viewH = RenderSize.Height;

            double right = HorizontalOffset + viewW;
            double bottom = VerticalOffset + viewH;

            int firstRow = FindIndex(_rowOffsets, VerticalOffset);
            int lastRow = Math.Min(_model.Rows - 1, FindIndex(_rowOffsets, bottom));

            int firstCol = FindIndex(_colOffsets, HorizontalOffset);
            int lastCol = Math.Min(_model.Columns - 1, FindIndex(_colOffsets, right));

            double y = _rowOffsets[firstRow] - VerticalOffset;

            for (int r = firstRow; r <= lastRow; r++)
            {
                double x = _colOffsets[firstCol] - HorizontalOffset;
                double rh = _model.RowHeights[r];

                for (int c = firstCol; c <= lastCol; c++)
                {
                    double cw = _model.ColumnWidths[c];
                    var cell = _model[r, c];

                    Rect rect = new Rect(x, y, cw, rh);
                    dc.DrawRectangle(cell.Background ?? Brushes.White, null, rect);

                    if (_searchHighlights.Contains((r, c)))
                        dc.DrawRectangle(new SolidColorBrush(Color.FromArgb(80, 255, 255, 0)), null, rect);

                    if (_model.ShowGridLines)
                        dc.DrawRectangle(null, new Pen(Brushes.Gray, 0.5), rect);

                    x += cw;
                }

                y += rh;
            }

            // rysowanie nagłówków tylko jeśli model istnieje
            _textRenderer?.Render(dc);
        }

        internal (int row, int col, Rect rect) HitTestCell(Point p)
        {
            if (_rowOffsets == null || _colOffsets == null) return (-1, -1, Rect.Empty);

            double px = p.X + HorizontalOffset;
            double py = p.Y + VerticalOffset;

            int row = FindIndex(_rowOffsets, py);
            int col = FindIndex(_colOffsets, px);

            if (row < 0 || col < 0 || row >= _model.Rows || col >= _model.Columns)
                return (-1, -1, Rect.Empty);

            Rect rect = new Rect(
                _colOffsets[col] - HorizontalOffset,
                _rowOffsets[row] - VerticalOffset,
                _model.ColumnWidths[col],
                _model.RowHeights[row]);

            return (row, col, rect);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (_model == null) return;

            var pos = e.GetPosition(this);
            var hit = HitTestCell(pos);

            if (hit.row >= 0 && hit.col == HyperlinkColumn)
            {
                var cell = _model[hit.row, HyperlinkColumn];
                if (cell.Hyperlink != null)
                {
                    Mouse.OverrideCursor = Cursors.Hand;
                    return;
                }
            }

            Mouse.OverrideCursor = null;
        }

        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            if (_model == null) return;

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
                _editorController.BeginEdit(hit.row, hit.col, hit.rect);

            InvalidateVisual();
            e.Handled = true;
        }

        public void ScrollToCell(int row, int col)
        {
            if (_model == null) return;

            double y = _rowOffsets[row];
            double x = _colOffsets[col];

            var sv = FindScrollViewer();
            sv?.ScrollToVerticalOffset(y);
            sv?.ScrollToHorizontalOffset(x);

            Rect rect = new Rect(
                x - HorizontalOffset,
                y - VerticalOffset,
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

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
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
