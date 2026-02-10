using System.Windows.Media;

namespace ExcelViewerV2.Model
{
    public enum HeaderFooterPosition
    {
        Left,
        Center,
        Right,
        Unknown
    }

    public enum WatermarkType
    {
        HeaderFooter,   // Z nagłówka/stopki
        Drawing,        // Obraz w arkuszu
        Background      // Tło arkusza
    }

    public class WatermarkImage
    {
        public ImageSource Image { get; set; }
        public HeaderFooterPosition Position { get; set; } = HeaderFooterPosition.Unknown;
        public WatermarkType Type { get; set; } = WatermarkType.HeaderFooter;
    }
}
