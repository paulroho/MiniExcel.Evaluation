using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets;

public class SpreadsheetWriter
{
    public void Write(string fileName, IEnumerable<Line> lines)
    {
        MiniExcel.SaveAs(fileName, lines);
    }
}