using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets;

public class SpreadsheetWriter
{
    public void WriteSpreadsheet(string fileName, IEnumerable<Line> lines)
    {
        MiniExcel.SaveAs(fileName, lines);
    }

    public void WriteSpreadsheet(string fileName, Dictionary<string, object> sheets)
    {
        MiniExcel.SaveAs(fileName, sheets);
    }
}