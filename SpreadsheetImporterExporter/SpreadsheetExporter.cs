using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets;

public class SpreadsheetExporter
{
    public void Write(string fileName, IEnumerable<Line> lines)
    {
        MiniExcel.SaveAs(fileName, lines);
    }
}