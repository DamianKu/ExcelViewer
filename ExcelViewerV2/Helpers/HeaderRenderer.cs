using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
using ExcelViewerV2.Controls;

namespace ExcelViewerV2.Helpers
{
    public static class HeaderRenderer
    {
        // =============================
        // COLUMN HEADERS
        // =============================
        public static void RefreshColumnHeaders(
            ExcelLikeGrid grid,
            Panel panel)
        {
            if (grid?.Model == null) return;

            panel.Children.Clear();

            double scrollX = grid.HorizontalOffset;
            double viewW = grid.RenderSize.Width;

            double acc = 0;

            for (int c = 0; c < grid.Model.Columns; c++)
            {
                double w = grid.Model.ColumnWidths[c];

                double start = acc;
                double end = acc + w;

                acc += w;

                if (end < scrollX) continue;
                if (start > scrollX + viewW) break;

                var tb = new TextBlock
                {
                    Text = GetExcelColumnName(c),
                    Width = w,
                    Height = 22,
                    TextAlignment = TextAlignment.Center,
                    FontWeight = FontWeights.Bold,
                    Background = Brushes.LightGray
                };

                Canvas.SetLeft(tb, start - scrollX);
                Canvas.SetTop(tb, 0);

                panel.Children.Add(tb);
            }
        }

        // =============================
        // ROW HEADERS
        // =============================
        public static void RefreshRowHeaders(
            ExcelLikeGrid grid,
            Panel panel)
        {
            if (grid?.Model == null) return;

            panel.Children.Clear();

            double scrollY = grid.VerticalOffset;
            double viewH = grid.RenderSize.Height;

            double acc = 0;

            for (int r = 0; r < grid.Model.Rows; r++)
            {
                double h = grid.Model.RowHeights[r];

                double start = acc;
                double end = acc + h;

                acc += h;

                if (end < scrollY) continue;
                if (start > scrollY + viewH) break;

                var tb = new TextBlock
                {
                    Text = (r + 1).ToString(),
                    Width = 40,
                    Height = h,
                    TextAlignment = TextAlignment.Center,
                    FontWeight = FontWeights.Bold,
                    Background = Brushes.LightGray
                };

                Canvas.SetTop(tb, start - scrollY);
                Canvas.SetLeft(tb, 0);

                panel.Children.Add(tb);
            }
        }

        // =============================
        // COLUMN NAME
        // =============================
        private static string GetExcelColumnName(int index)
        {
            string name = "";

            do
            {
                name = (char)('A' + (index % 26)) + name;
                index = (index / 26) - 1;
            }
            while (index >= 0);

            return name;
        }
    }
}
