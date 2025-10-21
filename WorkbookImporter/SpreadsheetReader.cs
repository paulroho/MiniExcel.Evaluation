using MiniExcelLibs;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks;

public class SpreadsheetReader
{
    public List<T> ReadSpreadsheet<T>(string fileName) 
        where T : class, new()
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<T>().ToList();
    }

    public List<T> ReadSpreadsheet<T>(string fileName, string sheetName, string startingCell = "A1") 
        where T : class, IRowMarker, new()
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<T>(sheetName, startCell: startingCell)
            .TakeWhile(row => !string.IsNullOrWhiteSpace(row.Marker))
            .ToList();
    }

    public interface IRowMarker
    {
        string? Marker { get; }
    }
}