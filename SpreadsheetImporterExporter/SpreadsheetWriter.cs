using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets;

public class SpreadsheetWriter
{
    public void WriteSpreadsheet(string fileName, IEnumerable<Line> lines)
    {
        MiniExcel.SaveAs(fileName, lines);
    }
}