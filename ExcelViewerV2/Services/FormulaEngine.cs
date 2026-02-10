using ExcelViewerV2.Model;
using ExcelViewerV2.Formula;

namespace ExcelViewerV2.Services
{
    public class FormulaEngine
    {
        private readonly SpreadsheetModel _sheet;

        public FormulaEngine(SpreadsheetModel sheet)
        {
            _sheet = sheet;
        }

        public void RecalculateAll()
        {
            for (int r = 0; r < _sheet.Rows; r++)
            for (int c = 0; c < _sheet.Columns; c++)
                EvaluateCell(r, c);
        }

        private void EvaluateCell(int row, int col)
        {
            var cell = _sheet[row, col];

            if (!cell.IsFormula)
            {
                cell.DisplayValue = cell.RawContent;
                return;
            }

            string formula = cell.RawContent[1..];

            try
            {
                var tokens = FormulaLexer.Tokenize(formula);

                var parser = new FormulaParser();
                var ast = parser.Parse(tokens);

                var evaluator = new FormulaEvaluator(_sheet);
                double val = evaluator.Eval(ast);

                cell.DisplayValue = val.ToString();
            }
            catch
            {
                cell.DisplayValue = "#ERR";
            }
        }
    }
}
