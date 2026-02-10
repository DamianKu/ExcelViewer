using ExcelViewerV2.Model;

namespace ExcelViewerV2.UndoRedo
{
    public sealed class EditSession
    {
        public CellModel Cell { get; }
        public string OriginalValue { get; }
        public string CurrentValue { get; private set; }

        public EditSession(CellModel cell)
        {
            Cell = cell;
            OriginalValue = cell.Text ?? "";
            CurrentValue = OriginalValue;
        }

        public void Update(string value)
        {
            CurrentValue = value ?? "";
        }

        public bool HasChanged => OriginalValue != CurrentValue;
    }
}
