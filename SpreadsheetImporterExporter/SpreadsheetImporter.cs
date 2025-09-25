using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets;

public class SpreadsheetImporter
{
    public List<Line> ImportSpreadsheet(string fileName)
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<Line>().ToList();
    }
}