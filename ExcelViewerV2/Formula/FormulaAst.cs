namespace ExcelViewerV2.Formula
{
    public abstract class AstNode { }

    public sealed class NumberNode : AstNode
    {
        public double Value;
        public NumberNode(double v) => Value = v;
    }

    public sealed class CellNode : AstNode
    {
        public string Address;
        public CellNode(string a) => Address = a;
    }

    public sealed class BinaryNode : AstNode
    {
        public string Op;
        public AstNode Left, Right;

        public BinaryNode(string op, AstNode l, AstNode r)
        {
            Op = op; Left = l; Right = r;
        }
    }

    public sealed class FunctionNode : AstNode
    {
        public string Name;
        public AstNode Arg1, Arg2;

        public FunctionNode(string n, AstNode a1, AstNode a2 = null)
        {
            Name = n; Arg1 = a1; Arg2 = a2;
        }
    }

    public sealed class RangeNode : AstNode
    {
        public string A, B;
        public RangeNode(string a, string b) { A = a; B = b; }
    }
}
