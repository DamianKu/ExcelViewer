using System.Windows;
using System.Windows.Controls;
using ExcelViewerV2.Editing;
using ExcelViewerV2.Model;
using ExcelViewerV2.Core.UndoRedo;

namespace ExcelViewerV2.Editing
{
    public sealed class CellEditorLayer
    {
        private readonly Canvas _canvas;
        private readonly CellEditorController _controller;

        public UIElement Visual => _canvas;

        // ✅ POPRAWIONY KONSTRUKTOR
        public CellEditorLayer(FrameworkElement owner, UndoManager undoManager)
        {
            _canvas = new Canvas
            {
                IsHitTestVisible = false
            };

            _controller = new CellEditorController(owner, undoManager);
            _canvas.Children.Add(_controller.EditorVisual);
        }

        public void AttachModel(SpreadsheetModel model)
        {
            _controller.AttachModel(model);
        }

        /// <summary>
        /// Rozpoczyna edycję komórki
        /// </summary>
        public void BeginEdit(int row, int column, Rect cellRect)
        {
            _controller.BeginEdit(row, column, cellRect);
        }

        /// <summary>
        /// Wywoływane z ArrangeOverride w ExcelLikeGrid
        /// </summary>
        public void ArrangeEditor()
        {
            _controller.ArrangeEditor();
        }

        /// <summary>
        /// Czy aktualnie trwa edycja
        /// </summary>
        public bool IsEditing => _controller != null;
    }
}
