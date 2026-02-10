using System;
using System.Windows;
using System.Windows.Media;

namespace ExcelViewerV2.Model
{
    public sealed class CellModel
    {
        // =========================
        // FORMUŁY – NOWE
        // =========================

        /// <summary>
        /// To wpisuje użytkownik (np. "=A1+B1" albo "123")
        /// </summary>
        public string RawContent { get; set; } = "";

        /// <summary>
        /// Wynik obliczeń (to renderujemy)
        /// </summary>
        public string DisplayValue { get; set; } = "";

        public bool IsFormula =>
            !string.IsNullOrEmpty(RawContent) &&
            RawContent.StartsWith("=");

        // =========================
        // TEKST (kompatybilność z rendererem)
        // =========================

        public string Text
        {
            get => DisplayValue;
            set
            {
                RawContent = value ?? "";
                DisplayValue = value ?? "";
            }
        }

        // =========================
        // CZCIONKA – EXCEL 1:1
        // =========================

        public string FontFamilyName { get; set; } = "Calibri";
        public double FontSize { get; set; } = 11;

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

        public bool IsEmpty =>
            string.IsNullOrEmpty(RawContent);

        public bool CanOverflow =>
            Wrapping == TextWrapping.NoWrap &&
            !string.IsNullOrEmpty(DisplayValue);

        // =========================
        // HIPERLINK
        // =========================

        public Uri Hyperlink { get; set; }
        public bool HasHyperlink => Hyperlink != null;

        // =========================
        // UNDO/REDO – WAŻNE
        // =========================

        public void SetTextInternal(string newText)
        {
            RawContent = newText ?? "";
            DisplayValue = newText ?? "";
        }
    }
}
