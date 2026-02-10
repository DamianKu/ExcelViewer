using System.Globalization;
using System.Windows;
using System.Windows.Media;
using ExcelViewerV2.Model;
using ExcelViewerV2.Controls;

namespace ExcelViewerV2.Rendering
{
    public sealed class CellTextRenderer
    {
        private const double CellPadding = 2;

        private readonly SpreadsheetModel _model;
        private readonly ExcelLikeGrid _grid;

        public CellTextRenderer(SpreadsheetModel model, Visual owner)
        {
            _model = model;
            _grid = owner as ExcelLikeGrid;
        }

        public void Render(DrawingContext dc)
        {
            if (_model == null) return;

            double dpi = _grid != null
                ? VisualTreeHelper.GetDpi(_grid).PixelsPerDip
                : 1.0;

            if (dpi <= 0) dpi = 1.0;

            // 🔥 KLUCZOWA POPRAWKA — translacja o scroll
            if (_grid != null)
                dc.PushTransform(
                    new TranslateTransform(
                        -_grid.HorizontalOffset,
                        -_grid.VerticalOffset));

            double y = 0;

            for (int r = 0; r < _model.Rows; r++)
            {
                double x = 0;
                double rh = _model.RowHeights[r];

                for (int c = 0; c < _model.Columns; c++)
                {
                    double cw = _model.ColumnWidths[c];
                    var cell = _model[r, c];

                    if (string.IsNullOrEmpty(cell.Text))
                    {
                        x += cw;
                        continue;
                    }

                    double availableWidth = cw - 2 * CellPadding;

                    if (cell.Wrapping == TextWrapping.NoWrap)
                    {
                        int cc = c + 1;
                        while (cc < _model.Columns &&
                               string.IsNullOrEmpty(_model[r, cc].Text))
                        {
                            availableWidth += _model.ColumnWidths[cc];
                            cc++;
                        }
                    }

                    var ft = new FormattedText(
                        cell.Text,
                        CultureInfo.CurrentCulture,
                        FlowDirection.LeftToRight,
                        new Typeface(
                            new FontFamily(cell.FontFamilyName ?? "Segoe UI"),
                            FontStyles.Normal,
                            cell.Bold ? FontWeights.Bold : FontWeights.Normal,
                            FontStretches.Normal),
                        cell.FontSize > 0 ? cell.FontSize : 12,
                        cell.Foreground ?? Brushes.Black,
                        dpi)
                    {
                        TextAlignment = TextAlignment.Left,
                        Trimming = TextTrimming.None,
                        MaxTextWidth = availableWidth,
                        MaxLineCount = cell.Wrapping == TextWrapping.Wrap
                            ? int.MaxValue
                            : 1
                    };

                    if (cell.Wrapping == TextWrapping.Wrap &&
                        ft.Height + 2 * CellPadding > rh)
                    {
                        rh = ft.Height + 2 * CellPadding;
                        _model.RowHeights[r] = rh;
                    }

                    double textX = x + CellPadding;
                    double textY = y + CellPadding;

                    switch (cell.VerticalAlignment)
                    {
                        case VerticalAlignment.Center:
                            textY = y + (rh - ft.Height) / 2;
                            break;
                        case VerticalAlignment.Bottom:
                            textY = y + rh - ft.Height - CellPadding;
                            break;
                    }

                    switch (cell.Alignment)
                    {
                        case TextAlignment.Center:
                            textX = x + (cw - ft.Width) / 2;
                            break;
                        case TextAlignment.Right:
                            textX = x + cw - ft.Width - CellPadding;
                            break;
                    }

                    if (cell.TextRotation != 0)
                    {
                        double cx = x + cw / 2;
                        double cy = y + rh / 2;

                        dc.PushTransform(
                            new RotateTransform(cell.TextRotation, cx, cy));

                        dc.DrawText(ft, new Point(textX, textY));

                        dc.Pop();
                    }
                    else
                    {
                        dc.DrawText(ft, new Point(textX, textY));
                    }

                    x += cw;
                }

                y += rh;
            }

            if (_grid != null)
                dc.Pop();
        }
    }
}
