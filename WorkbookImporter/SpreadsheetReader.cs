using MiniExcelLibs;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks;

public class SpreadsheetReader
{
    public List<Line> ReadSpreadsheet(string fileName)
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<Line>().ToList();
    }

    public List<Line> ReadSpreadsheet(string fileName, string sheetName, string startingCell = "A1")
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<Line>(sheetName, startCell: startingCell).ToList();
    }
}