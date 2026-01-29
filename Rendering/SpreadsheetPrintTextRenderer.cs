using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using ExcelViewerV2.Model;

namespace ExcelViewerV2.Rendering
{
    public sealed class SpreadsheetPrintTextRenderer
    {
        private const double CellPadding = 2.0;
        private readonly SpreadsheetModel _model;
        private readonly Dictionary<string, Typeface> _typefaceCache = new();

        public SpreadsheetPrintTextRenderer(SpreadsheetModel model)
        {
            _model = model;
        }

        public void Render(DrawingContext dc, double dpi, bool isPrinting)
        {
            if (_model == null) return;

            // =========================
            // HIPERLINKI – reset hit test
            // =========================
            _model.Hyperlinks.Clear();

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
                        GetTypeface(cell.FontFamilyName, cell.Bold),
                        cell.FontSize > 0 ? cell.FontSize : 12,
                        cell.HasHyperlink
                            ? Brushes.Blue
                            : (cell.Foreground ?? Brushes.Black),
                        dpi / 96.0)
                    {
                        MaxLineCount = cell.Wrapping == TextWrapping.Wrap ? int.MaxValue : 1,
                        MaxTextWidth = availableWidth,
                        TextAlignment = cell.Alignment,
                        Trimming = TextTrimming.None
                    };

                    // =========================
                    // HIPERLINK – styl Excela
                    // =========================
                    if (cell.HasHyperlink)
                        ft.SetTextDecorations(TextDecorations.Underline);

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

                    // =========================
                    // HIPERLINK – obszar kliku
                    // (TYLKO podgląd, NIE druk)
                    // =========================
                    if (!isPrinting && cell.HasHyperlink)
                    {
                        _model.Hyperlinks.Add(new HyperlinkHit
                        {
                            Rect = new Rect(textX, textY, ft.Width, ft.Height),
                            Link = cell.Hyperlink
                        });
                    }

                    x += cw;
                }

                y += rh;
            }
        }

        private Typeface GetTypeface(string fontFamilyName, bool bold)
        {
            if (string.IsNullOrWhiteSpace(fontFamilyName))
                fontFamilyName = "Calibri";

            string key = fontFamilyName + "|" + bold;
            if (_typefaceCache.TryGetValue(key, out var cached))
                return cached;

            FontFamily ff;
            try
            {
                ff = new FontFamily(fontFamilyName);
                var tfTest = new Typeface(
                    ff,
                    FontStyles.Normal,
                    bold ? FontWeights.Bold : FontWeights.Normal,
                    FontStretches.Normal);

                if (!tfTest.TryGetGlyphTypeface(out _))
                    ff = new FontFamily("Calibri");
            }
            catch
            {
                ff = new FontFamily("Calibri");
            }

            var tf = new Typeface(
                ff,
                FontStyles.Normal,
                bold ? FontWeights.Bold : FontWeights.Normal,
                FontStretches.Normal);

            _typefaceCache[key] = tf;
            return tf;
        }
    }
}
