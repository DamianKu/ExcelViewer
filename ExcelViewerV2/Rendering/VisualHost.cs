using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;

namespace ExcelViewerV2.Rendering
{
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
