
using ClosedXML.Excel;
using ExcelViewerV2.Model;


public static class ExcelSaveAdapter
{
public static void Save(string path, SpreadsheetModel model)
{
using var wb = new XLWorkbook();
var ws = wb.AddWorksheet("Sheet1");


for (int r = 0; r < model.Rows; r++)
for (int c = 0; c < model.Columns; c++)
ws.Cell(r + 1, c + 1).Value = model.Cells[r, c].Text;


wb.SaveAs(path);
}
}