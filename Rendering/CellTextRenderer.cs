using System.Globalization;
using System.Windows;
using System.Windows.Media;
using ExcelViewerV2.Model;

namespace ExcelViewerV2.Rendering
{
    public sealed class CellTextRenderer
    {
        private const double CellPadding = 2;

        private readonly SpreadsheetModel _model;
        private readonly Visual _owner;

        public CellTextRenderer(SpreadsheetModel model, Visual owner)
        {
            _model = model;
            _owner = owner;
        }

        public void Render(DrawingContext dc)
        {
            if (_model == null) return;

            double y = 0;
            double dpi = _owner != null ? VisualTreeHelper.GetDpi(_owner).PixelsPerDip : 1.0;

            for (int r = 0; r < _model.Rows; r++)
            {
                double x = 0;
                double rh = _model.RowHeights[r];

                for (int c = 0; c < _model.Columns; c++)
                {
                    double cw = _model.ColumnWidths[c];
                    var cell = _model[r, c];

                    if (!string.IsNullOrEmpty(cell.Text))
                    {
                        double availableWidth = cw - 2 * CellPadding;

                        if (cell.Wrapping == TextWrapping.NoWrap)
                        {
                            int cc = c + 1;
                            while (cc < _model.Columns && string.IsNullOrEmpty(_model[r, cc].Text))
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
                            MaxLineCount = cell.Wrapping == TextWrapping.Wrap ? int.MaxValue : 1,
                            MaxTextWidth = availableWidth,
                            TextAlignment = cell.Alignment,
                            Trimming = TextTrimming.None
                        };

                        double textX = x + CellPadding;
                        double textY = y + CellPadding;

                        if (cell.VerticalAlignment == VerticalAlignment.Center)
                            textY = y + (rh - ft.Height) / 2;
                        else if (cell.VerticalAlignment == VerticalAlignment.Bottom)
                            textY = y + rh - ft.Height - CellPadding;

                        if (cell.Alignment == TextAlignment.Center)
                            textX = x + (cw - ft.Width) / 2;
                        else if (cell.Alignment == TextAlignment.Right)
                            textX = x + cw - ft.Width - CellPadding;

                        dc.DrawText(ft, new Point(textX, textY));
                    }

                    x += cw;
                }

                y += rh;
            }
        }
    }
}
