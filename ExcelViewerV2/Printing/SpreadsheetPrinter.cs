using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using ExcelViewerV2.Model;
using ExcelViewerV2.Rendering;
using System.Windows.Markup;

namespace ExcelViewerV2.Printing
{
    public sealed class SpreadsheetPrinter
    {
        private readonly SpreadsheetRenderer _renderer = new SpreadsheetRenderer();

        // Marginesy w calach * 96 DPI
        private const double MarginLeft = 0.7 * 96;
        private const double MarginRight = 0.7 * 96;
        private const double MarginTop = 0.75 * 96;
        private const double MarginBottom = 0.75 * 96;

        // Ścieżka do raportu diagnostycznego
        public string DebugReportPath { get; set; } = "SpreadsheetPrintReport.txt";

        public void Print(SpreadsheetModel model)
        {
            if (model == null) throw new ArgumentNullException(nameof(model));

            try
            {
                var printDialog = new PrintDialog();
                if (printDialog.ShowDialog() != true) return;

                double pageWidth = printDialog.PrintableAreaWidth;
                double pageHeight = printDialog.PrintableAreaHeight;

                if (pageWidth <= 0 || pageHeight <= 0)
                    throw new InvalidOperationException("Nieprawidłowe wymiary strony do drukowania.");

                double usableWidth = pageWidth - MarginLeft - MarginRight;
                double usableHeight = pageHeight - MarginTop - MarginBottom;

                var fixedDoc = new FixedDocument();
                int currentRow = 0;

                LogReport($"--- Spreadsheet Print Report --- {DateTime.Now}");
                LogReport($"PageWidth: {pageWidth}, PageHeight: {pageHeight}");
                LogReport($"UsableWidth: {usableWidth}, UsableHeight: {usableHeight}");
                LogReport($"Total Rows: {model.Rows}, Total Columns: {model.Columns}");
                LogReport("");

                while (currentRow < model.Rows)
                {
                    double yAccum = 0;
                    int lastRowOnPage = currentRow;

                    while (lastRowOnPage < model.Rows &&
                           yAccum + model.RowHeights[lastRowOnPage] <= usableHeight)
                    {
                        yAccum += model.RowHeights[lastRowOnPage];
                        lastRowOnPage++;
                    }

                    if (lastRowOnPage == currentRow) lastRowOnPage++; // minimum 1 wiersz

                    var pageModel = CreateSubModel(model, currentRow, lastRowOnPage);

                    AddPage(pageModel, fixedDoc, pageWidth, pageHeight);

                    LogReport($"Page Rows: {currentRow} - {lastRowOnPage - 1}");
                    LogReport($"Row Heights: {string.Join(",", pageModel.RowHeights)}");
                    LogReport($"Column Widths: {string.Join(",", pageModel.ColumnWidths)}");
                    LogReport("");

                    currentRow = lastRowOnPage;
                }

                printDialog.PrintDocument(fixedDoc.DocumentPaginator, "Spreadsheet");
            }
            catch (Exception ex)
            {
                // zapis do raportu crashy
                File.AppendAllText(DebugReportPath,
                    $"[CRASH] {DateTime.Now}\n{ex}\n\n");

                // opcjonalnie pokaz alert
                MessageBox.Show(
                    $"Błąd podczas drukowania: {ex.Message}\nSzczegóły zapisano w {DebugReportPath}",
                    "Błąd drukowania", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ================================
        // Tworzenie submodelu dla pojedynczej strony
        private SpreadsheetModel CreateSubModel(SpreadsheetModel model, int startRow, int endRow)
        {
            int rowCount = endRow - startRow;
            var subModel = new SpreadsheetModel(rowCount, model.Columns);

            for (int c = 0; c < model.Columns; c++)
                subModel.ColumnWidths[c] = model.ColumnWidths[c];

            for (int r = 0; r < rowCount; r++)
                subModel.RowHeights[r] = model.RowHeights[startRow + r];

            for (int r = 0; r < rowCount; r++)
                for (int c = 0; c < model.Columns; c++)
                    subModel.Cells[r, c] = model.Cells[startRow + r, c];

            subModel.WatermarkImages.AddRange(model.WatermarkImages);
            subModel.HeaderImages.AddRange(model.HeaderImages);
            subModel.FooterImages.AddRange(model.FooterImages);

            subModel.ShowGridLines = model.ShowGridLines;

            return subModel;
        }

        private void AddPage(SpreadsheetModel pageModel, FixedDocument fixedDoc,
            double pageWidth, double pageHeight)
        {
            if (pageModel == null) return;

            var pageContent = new PageContent();
            var fixedPage = new FixedPage
            {
                Width = pageWidth,
                Height = pageHeight
            };

            double totalColumnsWidth = pageModel.ColumnWidths.Sum();
            double scaleX = totalColumnsWidth > 0
                ? (pageWidth - MarginLeft - MarginRight) / totalColumnsWidth
                : 1.0;

            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                dc.PushTransform(new TranslateTransform(MarginLeft, MarginTop));
                dc.PushTransform(new ScaleTransform(scaleX, 1.0));

                var renderer = new SpreadsheetRenderer();
                renderer.Render(dc, pageModel, dpi: 96.0, isPrinting: true);

                dc.Pop(); // scale
                dc.Pop(); // translate
            }

            var canvas = new Canvas
            {
                Width = pageWidth,
                Height = pageHeight
            };
            canvas.Children.Add(new VisualHost(visual));
            fixedPage.Children.Add(canvas);

            fixedPage.Measure(new Size(pageWidth, pageHeight));
            fixedPage.Arrange(new Rect(0, 0, pageWidth, pageHeight));
            fixedPage.UpdateLayout();

            ((IAddChild)pageContent).AddChild(fixedPage);
            fixedDoc.Pages.Add(pageContent);
        }

        // ========================
        // zapis raportu tekstowego
        private void LogReport(string text)
        {
            if (string.IsNullOrEmpty(DebugReportPath)) return;
            try
            {
                File.AppendAllText(DebugReportPath, text + Environment.NewLine);
            }
            catch { /* ignoruj błędy zapisu */ }
        }
    }

    // =======================
    public class VisualHost : FrameworkElement
    {
        private readonly Visual _visual;
        public VisualHost(Visual visual) => _visual = visual;
        protected override int VisualChildrenCount => 1;
        protected override Visual GetVisualChild(int index) => _visual;
    }
}
