using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using ExcelViewerV2.Services;

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
        // FORMUŁY – NOWE
        // =========================

        public FormulaEngine FormulaEngine { get; }

        // =========================
        // SIATKA / WATERMARKI
        // =========================

        public bool ShowGridLines { get; set; } = true;

        public List<WatermarkImage> WatermarkImages { get; } = new();
        public List<WatermarkImage> HeaderImages { get; } = new();
        public List<WatermarkImage> FooterImages { get; } = new();

        // =========================
        // UKRYWANIE ARKUSZY
        // =========================

        public bool IsVisible { get; set; } = true;

        // =========================
        // HIPERLINKI – HIT TEST
        // =========================

        public List<HyperlinkHit> Hyperlinks { get; } = new();

        // =========================
        // KONSTRUKTOR
        // =========================

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

            // inicjalizacja silnika formuł
            FormulaEngine = new FormulaEngine(this);
        }

        // =========================
        // INDEKSER
        // =========================

        public CellModel this[int row, int col] => Cells[row, col];

        // =========================
        // USTAWIANIE WARTOŚCI KOMÓRKI
        // (WAŻNE – używać zamiast cell.Text=)
        // =========================

        public void SetCell(int row, int col, string value)
        {
            if (row < 0 || col < 0 || row >= Rows || col >= Columns)
                return;

            var cell = Cells[row, col];

            cell.RawContent = value ?? "";
            cell.DisplayValue = value ?? "";

            // przelicz formuły
            FormulaEngine.RecalculateAll();
        }

        // =========================
        // RĘCZNE PRZELICZENIE
        // =========================

        public void Recalculate()
        {
            FormulaEngine.RecalculateAll();
        }

        // =========================
        // ITERACJA PO KOMÓRKACH
        // =========================

        public IEnumerable<(int row, int col, CellModel cell)> EnumerateCells()
        {
            for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Columns; c++)
                yield return (r, c, Cells[r, c]);
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
