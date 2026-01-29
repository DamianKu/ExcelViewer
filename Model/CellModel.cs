using System.Windows;
using System.Windows.Media;

namespace ExcelViewerV2.Model
{
    public sealed class CellModel
    {
        // =========================
        // TEKST
        // =========================
        public string Text { get; set; }

        // =========================
        // CZCIONKA – EXCEL 1:1
        // =========================
        public string FontFamilyName { get; set; } = "Calibri";   // Excel default
        public double FontSize { get; set; } = 11;                // Excel default

        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikeout { get; set; }

        // =========================
        // KOLORY
        // =========================
        public Brush Foreground { get; set; } = Brushes.Black;
        public Brush Background { get; set; } = Brushes.White;

        // =========================
        // WYRÓWNANIA
        // =========================
        public TextAlignment Alignment { get; set; } = TextAlignment.Left;
        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Center;

        public TextWrapping Wrapping { get; set; } = TextWrapping.NoWrap;

        // =========================
        // ROTACJA
        // =========================
        public double TextRotation { get; set; } = 0.0;

        // =========================
        // OBRAMOWANIA
        // =========================
        public bool BorderLeft { get; set; }
        public bool BorderTop { get; set; }
        public bool BorderRight { get; set; }
        public bool BorderBottom { get; set; }

        public Brush BorderBrush { get; set; } = Brushes.Black;
        public double BorderThickness { get; set; } = 1.0;

        // =========================
        // POMOCNICZE
        // =========================

        /// <summary>
        /// Czy komórka jest wizualnie pusta
        /// </summary>
        public bool IsEmpty =>
            string.IsNullOrEmpty(Text);

        /// <summary>
        /// Czy tekst może robić Excel-style overflow
        /// </summary>
        public bool CanOverflow =>
            Wrapping == TextWrapping.NoWrap &&
            !string.IsNullOrEmpty(Text);


        // =========================
        // HIPERLINK
        // =========================
        public Uri Hyperlink { get; set; }
        public bool HasHyperlink => Hyperlink != null;

        // =========================
        // NOWA METODA DLA UNDO/REDO
        // =========================
        public void SetTextInternal(string newText)
        {
            Text = newText ?? "";
        }

    }
}
