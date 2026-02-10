using ExcelViewerV2.Model;
using ExcelViewerV2.UndoRedo;

namespace ExcelViewerV2.Commands
{
    public sealed class CellTextChangeCommand : IUndoableCommand
    {
        private readonly CellModel _cell;
        private readonly string _oldText;
        private readonly string _newText;

        public CellTextChangeCommand(CellModel cell, string oldText, string newText)
        {
            _cell = cell;
            _oldText = oldText;
            _newText = newText;
        }

        public void Execute() => _cell.SetTextInternal(_newText);
        public void Undo() => _cell.SetTextInternal(_oldText);
    }
}
