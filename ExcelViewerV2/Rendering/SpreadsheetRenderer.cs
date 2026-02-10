using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ExcelViewerV2.Model;

namespace ExcelViewerV2.Rendering
{
    public sealed class SpreadsheetRenderer
    {
        public void Render(DrawingContext dc, SpreadsheetModel model, double dpi, bool isPrinting)
        {
            if (model == null) return;

            RenderBackground(dc, model, isPrinting);
            RenderSheetWatermarks(dc, model, dpi, isPrinting);
            RenderHeaderFooter(dc, model.HeaderImages, model, dpi, isPrinting, true);
            RenderHeaderFooter(dc, model.FooterImages, model, dpi, isPrinting, false);

            var textRenderer = new SpreadsheetPrintTextRenderer(model);
            textRenderer.Render(dc, dpi, isPrinting);
        }

        private void RenderBackground(DrawingContext dc, SpreadsheetModel model, bool isPrinting)
        {
            double y = 0;
            for (int r = 0; r < model.Rows; r++)
            {
                double x = 0;
                double rh = model.RowHeights[r];

                for (int c = 0; c < model.Columns; c++)
                {
                    double cw = model.ColumnWidths[c];
                    var cell = model[r, c];

                    dc.DrawRectangle(cell.Background ?? Brushes.White, null, new Rect(x, y, cw, rh));

                    if (!isPrinting && model.ShowGridLines)
                        dc.DrawRectangle(null, new Pen(Brushes.LightGray, 0.5), new Rect(x, y, cw, rh));

                    x += cw;
                }
                y += rh;
            }
        }

        private void RenderSheetWatermarks(DrawingContext dc, SpreadsheetModel model, double dpi, bool isPrinting)
        {
            if (!model.WatermarkImages.Any()) return;

            double totalWidth = model.ColumnWidths.Sum();
            double totalHeight = model.RowHeights.Sum();

            foreach (var wm in model.WatermarkImages)
            {
                if (wm.Image is not BitmapSource bmp) continue;
                Rect targetRect = new Rect(0, 0, totalWidth, totalHeight);
                dc.PushOpacity(isPrinting ? 0.3 : 0.12);
                dc.DrawImage(bmp, targetRect);
                dc.Pop();
            }
        }

        private void RenderHeaderFooter(DrawingContext dc,
            System.Collections.Generic.List<WatermarkImage> images,
            SpreadsheetModel model,
            double dpi,
            bool isPrinting,
            bool isHeader)
        {
            if (!images.Any()) return;

            double totalWidth = model.ColumnWidths.Sum();
            double totalHeight = model.RowHeights.Sum();

            double yOffset = isHeader ? -totalHeight * 0.01 : totalHeight + totalHeight * 0.005;

            foreach (var wm in images)
            {
                if (wm.Image is not BitmapSource bmp) continue;

                double imgWidth = bmp.PixelWidth * 96.0 / bmp.DpiX;
                double imgHeight = bmp.PixelHeight * 96.0 / bmp.DpiY;

                double x = wm.Position switch
                {
                    HeaderFooterPosition.Left => totalWidth * 0.05,
                    HeaderFooterPosition.Center => (totalWidth - imgWidth) / 2,
                    HeaderFooterPosition.Right => totalWidth - imgWidth - totalWidth * 0.05,
                    _ => (totalWidth - imgWidth) / 2
                };

                dc.PushOpacity(isPrinting ? 0.1 : 0.05);
                dc.DrawImage(bmp, new Rect(x, yOffset, imgWidth, imgHeight));
                dc.Pop();
            }
        }
    }
}
