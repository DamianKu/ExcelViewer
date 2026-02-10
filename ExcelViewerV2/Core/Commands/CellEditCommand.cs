using ExcelViewerV2.Model;

namespace ExcelViewerV2.Core.Commands
{
    public sealed class CellEditCommand : IUndoableCommand
    {
        private readonly CellModel _cell;
        private readonly string _oldValue;
        private readonly string _newValue;

        public CellEditCommand(CellModel cell, string oldValue, string newValue)
        {
            _cell = cell;
            _oldValue = oldValue;
            _newValue = newValue;
        }

        public void Execute()
        {
            _cell.Text = _newValue;
        }

        public void Undo()
        {
            _cell.Text = _oldValue;
        }
    }
}
