using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
        private readonly string _debugReportPath;

        public SpreadsheetPrintTextRenderer(SpreadsheetModel model, string debugReportPath = null)
        {
            _model = model ?? throw new ArgumentNullException(nameof(model));
            _debugReportPath = debugReportPath;
        }

        public void Render(DrawingContext dc, double dpi, bool isPrinting)
        {
            if (_model == null) return;

            _model.Hyperlinks.Clear();

            double y = 0;

            using StreamWriter writer = !string.IsNullOrEmpty(_debugReportPath)
                ? new StreamWriter(_debugReportPath, true)
                : null;

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
                        try
                        {
                            DrawCellText(dc, cell, r, c, x, y, cw, rh, dpi, isPrinting, writer);
                        }
                        catch (Exception ex)
                        {
                            writer?.WriteLine($"[ERROR] Row {r}, Col {c}: {ex.Message}");
                        }
                    }

                    x += cw;
                }

                y += rh;
            }
        }

        private void DrawCellText(
            DrawingContext dc,
            CellModel cell,
            int row,
            int col,
            double x,
            double y,
            double cw,
            double rh,
            double dpi,
            bool isPrinting,
            StreamWriter writer = null)
        {
            if (cell == null || string.IsNullOrEmpty(cell.Text)) return;

            double safeDpi = dpi > 0 ? dpi : 1.0;

            var typeface = GetTypeface(cell.FontFamilyName, cell.Bold);

            // ==============================
            // Obliczenie dostępnej szerokości z overflow
            // ==============================
            double availableWidth = cw - 2 * CellPadding;

            if (cell.CanOverflow)
            {
                int cc = col + 1;
                while (cc < _model.Columns && string.IsNullOrEmpty(_model[row, cc].Text))
                {
                    availableWidth += _model.ColumnWidths[cc];
                    cc++;
                }
            }

            // 🔒 Zabezpieczenie przed zerową, ujemną lub nieskończoną szerokością
            if (availableWidth <= 0 || double.IsInfinity(availableWidth) || double.IsNaN(availableWidth))
            {
                writer?.WriteLine($"[WARN] Row {row}, Col {col}: Invalid availableWidth={availableWidth}, replaced with 1000");
                availableWidth = 1000; // dowolna rozsądna szerokość dla drukowania
            }

            // ==============================
            // FormattedText z wrap i MaxLineCount
            // ==============================
            var ft = new FormattedText(
                cell.Text,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                typeface,
                cell.FontSize > 0 ? cell.FontSize : 12,
                cell.Foreground ?? Brushes.Black,
                safeDpi)
            {
                TextAlignment = cell.Alignment,
                Trimming = TextTrimming.None,
                MaxTextWidth = cell.Wrapping == TextWrapping.Wrap ? availableWidth : availableWidth,
                MaxLineCount = cell.Wrapping == TextWrapping.Wrap ? int.MaxValue : 1
            };

            if (cell.HasHyperlink)
                ft.SetTextDecorations(TextDecorations.Underline);

            // ==============================
            // Pozycjonowanie tekstu
            // ==============================
            double tx = x + CellPadding;
            double ty = y + CellPadding;

            // Pionowe centrowanie
            if (cell.VerticalAlignment == VerticalAlignment.Center)
                ty = y + (rh - ft.Height) / 2;
            else if (cell.VerticalAlignment == VerticalAlignment.Bottom)
                ty = y + rh - ft.Height - CellPadding;

            // Poziome centrowanie
            if (cell.Alignment == TextAlignment.Center)
                tx = x + (cw - ft.Width) / 2;
            else if (cell.Alignment == TextAlignment.Right)
                tx = x + cw - ft.Width - CellPadding;

            // ==============================
            // ROTACJA TEKSTU
            // ==============================
            if (Math.Abs(cell.TextRotation) < 0.01)
            {
                dc.DrawText(ft, new Point(tx, ty));
            }
            else
            {
                double centerX = x + cw / 2;
                double centerY = y + rh / 2;

                dc.PushTransform(new RotateTransform(cell.TextRotation, centerX, centerY));

                double rotatedTx = centerX - ft.Width / 2;
                double rotatedTy = centerY - ft.Height / 2;

                dc.DrawText(ft, new Point(rotatedTx, rotatedTy));

                dc.Pop();
            }

            // ==============================
            // Dodanie hyperlinka
            // ==============================
            if (!isPrinting && cell.HasHyperlink)
            {
                _model.Hyperlinks.Add(new HyperlinkHit
                {
                    Rect = new Rect(tx, ty, ft.Width, ft.Height),
                    Link = cell.Hyperlink
                });
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
                var test = new Typeface(
                    ff,
                    FontStyles.Normal,
                    bold ? FontWeights.Bold : FontWeights.Normal,
                    FontStretches.Normal);

                if (!test.TryGetGlyphTypeface(out _))
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
