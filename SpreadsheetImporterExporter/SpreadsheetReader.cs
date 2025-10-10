using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets;

public class SpreadsheetReader
{
    public List<Line> ReadSpreadsheet(string fileName)
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<Line>().ToList();
    }

    public List<Line> ReadSpreadsheet(string fileName, string sheetName)
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<Line>(sheetName).ToList();
    }
}