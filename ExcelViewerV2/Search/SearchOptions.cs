namespace ExcelViewerV2.Search
{
    public sealed class SearchOptions
    {
        public string Query { get; set; } = string.Empty;
        public bool MatchCase { get; set; }
        public bool MatchWholeCell { get; set; }
    }
}
