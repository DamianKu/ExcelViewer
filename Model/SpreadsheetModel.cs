using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace ExcelViewerV2.Model
{
    public sealed class SpreadsheetModel
    {
        // =========================
        // PODSTAWOWE INFO O ARKUSZU
        // =========================
        public string Name { get; set; } = "";

        public int Rows { get; }
        public int Columns { get; }

        // =========================
        // DANE KOMÓREK
        // =========================
        public CellModel[,] Cells { get; }
        public double[] ColumnWidths { get; }
        public double[] RowHeights { get; }

        // =========================
        // SIATKA / WATERMARKI
        // =========================
        public bool ShowGridLines { get; set; } = true;
        public List<WatermarkImage> WatermarkImages { get; } = new();
        public List<WatermarkImage> HeaderImages { get; } = new();
        public List<WatermarkImage> FooterImages { get; } = new();
        //==========================
        // UKRYWANIE ARKUSZY
        //==========================
        public bool IsVisible { get; set; } = true;

        // =========================
        // HIPERLINKI – HIT TEST
        // (renderowane dynamicznie)
        // =========================
        public List<HyperlinkHit> Hyperlinks { get; } = new();

        public SpreadsheetModel(int rows, int columns, string name = "")
        {
            Name = name;

            Rows = rows;
            Columns = columns;

            Cells = new CellModel[rows, columns];
            ColumnWidths = new double[columns];
            RowHeights = new double[rows];

            for (int r = 0; r < rows; r++)
            {
                RowHeights[r] = 22;

                for (int c = 0; c < columns; c++)
                {
                    if (r == 0)
                        ColumnWidths[c] = 100;

                    Cells[r, c] = new CellModel();
                }
            }
        }

        // =========================
        // INDEKSER
        // =========================
        public CellModel this[int row, int col] => Cells[row, col];

        // =========================
        // HELPER – iteracja po komórkach
        // (używane przez search)
        // =========================
        public IEnumerable<(int row, int col, CellModel cell)> EnumerateCells()
        {
            for (int r = 0; r < Rows; r++)
            {
                for (int c = 0; c < Columns; c++)
                {
                    yield return (r, c, Cells[r, c]);
                }
            }
        }
    }

    // =========================
    // HIPERLINK – OBSZAR KLIKNIĘCIA
    // =========================
    public sealed class HyperlinkHit
    {
        public Rect Rect { get; set; }
        public Uri Link { get; set; }
    }
}
