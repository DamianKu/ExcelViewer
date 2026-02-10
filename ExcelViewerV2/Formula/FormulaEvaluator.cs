using ExcelViewerV2.Model;

namespace ExcelViewerV2.Formula
{
    public sealed class FormulaEvaluator
    {
        private readonly SpreadsheetModel _sheet;

        public FormulaEvaluator(SpreadsheetModel sheet)
        {
            _sheet = sheet;
        }

        public double Eval(AstNode n)
        {
            switch (n)
            {
                case NumberNode num:
                    return num.Value;

                case CellNode cell:
                    var (r,c) = CellAddress.Parse(cell.Address);
                    var val = _sheet[r,c].DisplayValue;
                    return double.TryParse(val, out var d) ? d : 0;

                case BinaryNode bin:
                    var l = Eval(bin.Left);
                    var r2 = Eval(bin.Right);
                    return bin.Op switch
                    {
                        "+" => l + r2,
                        "-" => l - r2,
                        "*" => l * r2,
                        "/" => r2 == 0 ? 0 : l / r2,
                        _ => 0
                    };

                case RangeNode range:
                    return SumRange(range.A, range.B);

                case FunctionNode fn:
                    if (fn.Name == "SUM" && fn.Arg1 is RangeNode rn)
                        return SumRange(rn.A, rn.B);
                    break;
            }

            return 0;
        }

        private double SumRange(string a, string b)
        {
            var A = CellAddress.Parse(a);
            var B = CellAddress.Parse(b);

            double s = 0;

            for (int r=A.row; r<=B.row; r++)
            for (int c=A.col; c<=B.col; c++)
            {
                var v=_sheet[r,c].DisplayValue;
                if(double.TryParse(v,out var d)) s+=d;
            }

            return s;
        }
    }
}
