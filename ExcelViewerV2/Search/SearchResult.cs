namespace ExcelViewerV2.Search
{
    public sealed class SearchResult
    {
        public string SheetName { get; }
        public int Row { get; }
        public int Column { get; }
        public string Text { get; }

        public SearchResult(string sheetName, int row, int column, string text)
        {
            SheetName = sheetName;
            Row = row;
            Column = column;
            Text = text;
        }

        public override string ToString() => $"{SheetName}!{Row + 1}:{Column + 1} → {Text}";

    }
}
