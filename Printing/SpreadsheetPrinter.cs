using System;
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

        public void Print(SpreadsheetModel model)
        {
            if (model == null) throw new ArgumentNullException(nameof(model));

            var printDialog = new PrintDialog();
            if (printDialog.ShowDialog() != true) return;

            double pageWidth = printDialog.PrintableAreaWidth;
            double pageHeight = printDialog.PrintableAreaHeight;

            double usableWidth = pageWidth - MarginLeft - MarginRight;
            double usableHeight = pageHeight - MarginTop - MarginBottom;

            var fixedDoc = new FixedDocument();

            int currentRow = 0;

            while (currentRow < model.Rows)
            {
                double yAccum = 0;
                int lastRowOnPage = currentRow;

                // wybieramy ile wierszy zmieści się na stronie
                while (lastRowOnPage < model.Rows &&
                       yAccum + model.RowHeights[lastRowOnPage] <= usableHeight)
                {
                    yAccum += model.RowHeights[lastRowOnPage];
                    lastRowOnPage++;
                }

                // minimum 1 wiersz
                if (lastRowOnPage == currentRow) lastRowOnPage++;

                // Tworzymy submodel dla strony
                var pageModel = model.CreateSubModel(currentRow, lastRowOnPage);

                AddPage(pageModel, fixedDoc, pageWidth, pageHeight);

                currentRow = lastRowOnPage;
            }

            printDialog.PrintDocument(fixedDoc.DocumentPaginator, "Spreadsheet");
        }

        private void AddPage(SpreadsheetModel pageModel, FixedDocument fixedDoc,
            double pageWidth, double pageHeight)
        {
            var pageContent = new PageContent();
            var fixedPage = new FixedPage
            {
                Width = pageWidth,
                Height = pageHeight
            };

            // Obliczamy sumę szerokości kolumn strony
            double totalColumnsWidth = pageModel.ColumnWidths.Sum();
            double scaleX = (pageWidth - MarginLeft - MarginRight) / totalColumnsWidth;

            var visual = new DrawingVisual();
            using (var dc = visual.RenderOpen())
            {
                // przesunięcie marginesów
                dc.PushTransform(new TranslateTransform(MarginLeft, MarginTop));

                // skalujemy tylko w poziomie, by tło/watermarky nagłówki/stopki wypełniały szerokość
                dc.PushTransform(new ScaleTransform(scaleX, 1.0));

                _renderer.Render(dc, pageModel, dpi: 96.0, isPrinting: true);

                dc.Pop(); // koniec skalowania X
                dc.Pop(); // koniec przesunięcia marginesów
            }

            fixedPage.Children.Add(new VisualHost(visual));
            ((IAddChild)pageContent).AddChild(fixedPage);
            fixedDoc.Pages.Add(pageContent);
        }
    }

    // ===========================
    public static class SpreadsheetModelExtensions
    {
        public static SpreadsheetModel CreateSubModel(this SpreadsheetModel model, int startRow, int endRow)
        {
            int rowCount = endRow - startRow;
            var subModel = new SpreadsheetModel(rowCount, model.Columns);

            // ColumnWidths
            for (int c = 0; c < model.Columns; c++)
                subModel.ColumnWidths[c] = model.ColumnWidths[c];

            // RowHeights
            for (int r = 0; r < rowCount; r++)
                subModel.RowHeights[r] = model.RowHeights[startRow + r];

            // komórki
            for (int r = 0; r < rowCount; r++)
                for (int c = 0; c < model.Columns; c++)
                    subModel.Cells[r, c] = model.Cells[startRow + r, c];

            // watermark, nagłówki i stopki
            subModel.WatermarkImages.AddRange(model.WatermarkImages);
            subModel.HeaderImages.AddRange(model.HeaderImages);
            subModel.FooterImages.AddRange(model.FooterImages);

            // gridlines
            subModel.ShowGridLines = model.ShowGridLines;

            return subModel;
        }
    }

    // ===========================
    // Pomocniczy host do DrawingVisual
    public class VisualHost : FrameworkElement
    {
        private readonly Visual _visual;

        public VisualHost(Visual visual)
        {
            _visual = visual;
        }

        protected override int VisualChildrenCount => 1;

        protected override Visual GetVisualChild(int index) => _visual;
    }
}
