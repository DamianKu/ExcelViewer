using System.Collections.Generic;

namespace ExcelViewerV2.Formula
{
    public sealed class FormulaParser
    {
        private List<FormulaToken> _t;
        private int _i;

        public AstNode Parse(List<FormulaToken> tokens)
        {
            _t = tokens;
            _i = 0;
            return ParseExpr();
        }

        private FormulaToken Peek() => _t[_i];
        private FormulaToken Next() => _t[_i++];

        private AstNode ParseExpr()
        {
            var left = ParseTerm();

            while (Peek().Type == TokenType.Operator &&
                   (Peek().Text == "+" || Peek().Text == "-"))
            {
                var op = Next().Text;
                var right = ParseTerm();
                left = new BinaryNode(op, left, right);
            }

            return left;
        }

        private AstNode ParseTerm()
        {
            var left = ParseFactor();

            while (Peek().Type == TokenType.Operator &&
                   (Peek().Text == "*" || Peek().Text == "/"))
            {
                var op = Next().Text;
                var right = ParseFactor();
                left = new BinaryNode(op, left, right);
            }

            return left;
        }

        private AstNode ParseFactor()
        {
            var t = Next();

            if (t.Type == TokenType.Number)
                return new NumberNode(double.Parse(t.Text));

            if (t.Type == TokenType.Cell)
            {
                if (Peek().Type == TokenType.Colon)
                {
                    Next();
                    var b = Next();
                    return new RangeNode(t.Text, b.Text);
                }

                return new CellNode(t.Text);
            }

            if (t.Type == TokenType.Function)
            {
                Next(); // (
                var a1 = ParseExpr();

                AstNode a2 = null;
                if (Peek().Type == TokenType.Comma)
                {
                    Next();
                    a2 = ParseExpr();
                }

                Next(); // )
                return new FunctionNode(t.Text, a1, a2);
            }

            if (t.Type == TokenType.LParen)
            {
                var e = ParseExpr();
                Next();
                return e;
            }

            return new NumberNode(0);
        }
    }
}
