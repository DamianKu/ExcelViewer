using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ExcelViewerV2.Formula
{
    public static class FormulaLexer
    {
        static Regex number = new(@"^\d+(\.\d+)?");
        static Regex cell = new(@"^[A-Z]+\d+");
        static Regex func = new(@"^[A-Z]+(?=\()");

        public static List<FormulaToken> Tokenize(string s)
        {
            var tokens = new List<FormulaToken>();
            s = s.ToUpper();

            while (s.Length > 0)
            {
                if (char.IsWhiteSpace(s[0]))
                {
                    s = s[1..];
                    continue;
                }

                var mNum = number.Match(s);
                if (mNum.Success)
                {
                    tokens.Add(new(TokenType.Number, mNum.Value));
                    s = s[mNum.Length..];
                    continue;
                }

                var mCell = cell.Match(s);
                if (mCell.Success)
                {
                    tokens.Add(new(TokenType.Cell, mCell.Value));
                    s = s[mCell.Length..];
                    continue;
                }

                var mFunc = func.Match(s);
                if (mFunc.Success)
                {
                    tokens.Add(new(TokenType.Function, mFunc.Value));
                    s = s[mFunc.Length..];
                    continue;
                }

                char c = s[0];
                s = s[1..];

                switch (c)
                {
                    case '+':
                    case '-':
                    case '*':
                    case '/':
                        tokens.Add(new(TokenType.Operator, c.ToString())); break;
                    case '(':
                        tokens.Add(new(TokenType.LParen, "(")); break;
                    case ')':
                        tokens.Add(new(TokenType.RParen, ")")); break;
                    case ',':
                        tokens.Add(new(TokenType.Comma, ",")); break;
                    case ':':
                        tokens.Add(new(TokenType.Colon, ":")); break;
                }
            }

            tokens.Add(new(TokenType.End, ""));
            return tokens;
        }
    }
}
