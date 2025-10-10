using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets;

public class SpreadsheetWriter
{
    public void WriteSpreadsheet(string fileName, IEnumerable<Line> lines)
    {
        MiniExcel.SaveAs(fileName, lines);
    }

    public void WriteSpreadsheet(string fileName, Dictionary<string, Line[]> sheets)
    {
        var sheetData = sheets.ToDictionary(
            s => s.Key,
             object (s) => s.Value
        );
        MiniExcel.SaveAs(fileName, sheetData);
    }
}