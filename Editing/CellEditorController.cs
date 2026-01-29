using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ExcelViewerV2.Model;
using ExcelViewerV2.Core.UndoRedo;
using ExcelViewerV2.Commands;
using ExcelViewerV2.UndoRedo;

namespace ExcelViewerV2.Editing
{
    public sealed class CellEditorController
    {
        private const double CellPadding = 2;

        private readonly FrameworkElement _owner;
        private readonly TextBox _editor;
        private readonly UndoManager _undoManager;

        private SpreadsheetModel _model;

        private Rect _editorRect;
        private EditSession _editSession;

        public UIElement EditorVisual => _editor;

        public CellEditorController(FrameworkElement owner, UndoManager undoManager)
        {
            _owner = owner;
            _undoManager = undoManager ?? throw new ArgumentNullException(nameof(undoManager));

            _editor = new TextBox
            {
                Visibility = Visibility.Collapsed,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(CellPadding),
                AcceptsReturn = true,
                Background = Brushes.White,

                // 🔥 KLUCZOWE — WYŁĄCZAMY UNDO TEXTBOXA
                IsUndoEnabled = false,
                UndoLimit = 0
            };

            _editor.TextChanged += Editor_TextChanged;
            _editor.KeyDown += Editor_KeyDown;
            _editor.LostKeyboardFocus += (_, _) => CommitEdit();

            // MUSI BYĆ W DRZEWIE WIZUALNYM
            if (_owner is Panel panel)
                panel.Children.Add(_editor);
        }

        public void AttachModel(SpreadsheetModel model)
        {
            _model = model;
        }

        public void BeginEdit(int row, int col, Rect rect)
        {
            if (_model == null)
                return;

            var cell = _model[row, col];
            if (cell == null)
                return;

            _editorRect = rect;

            _editSession = new EditSession(cell);

            _editor.Text = cell.Text ?? string.Empty;
            _editor.FontSize = cell.FontSize > 0 ? cell.FontSize : 12;
            _editor.FontFamily = new FontFamily(cell.FontFamilyName ?? "Segoe UI");
            _editor.TextWrapping = cell.Wrapping;
            _editor.Visibility = Visibility.Visible;

            _owner.InvalidateArrange();

            _editor.Focus();
            _editor.SelectAll();
        }

        public void ArrangeEditor()
        {
            if (_editor.Visibility == Visibility.Visible)
                _editor.Arrange(_editorRect);
            else
                _editor.Arrange(new Rect(0, 0, 0, 0));
        }

        private void Editor_TextChanged(object sender, TextChangedEventArgs e)
        {
            _editSession?.Update(_editor.Text);
        }

        private void CommitEdit()
        {
            if (_editSession == null)
            {
                CleanupEdit();
                return;
            }

            if (_editSession.HasChanged)
            {
                _undoManager.Execute(
                    new CellTextChangeCommand(
                        _editSession.Cell,
                        _editSession.OriginalValue,
                        _editSession.CurrentValue));
            }

            CleanupEdit();
        }

        private void CancelEdit()
        {
            if (_editSession != null)
                _editSession.Cell.Text = _editSession.OriginalValue;

            CleanupEdit();
        }

        private void CleanupEdit()
        {
            _editSession = null;

            _editor.Visibility = Visibility.Collapsed;
            _owner.InvalidateVisual();
        }

        private void Editor_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitEdit();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                CancelEdit();
                e.Handled = true;
            }
        }

        public void EndEdit()
        {
            CommitEdit();
        }
    }
}
