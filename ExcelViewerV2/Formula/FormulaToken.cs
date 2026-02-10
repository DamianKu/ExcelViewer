namespace ExcelViewerV2.Formula
{
    public enum TokenType
    {
        Number,
        Operator,
        Cell,
        Function,
        LParen,
        RParen,
        Comma,
        Colon,
        End
    }

    public sealed class FormulaToken
    {
        public TokenType Type { get; }
        public string Text { get; }

        public FormulaToken(TokenType type, string text)
        {
            Type = type;
            Text = text;
        }
    }
}
