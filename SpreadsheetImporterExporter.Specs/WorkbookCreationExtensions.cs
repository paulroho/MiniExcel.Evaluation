using MiniExcelLibs;

namespace Evaluate.ReadingWritingSpreadsheets.Specs;

public static class WorkbookCreationExtensions
{
    public static TemporaryFile SaveAsWorkbook(this Dictionary<string, Line[]> sheets)
    {
        var file = new TemporaryFile(".xlsx");

        var sheetData = sheets.ToDictionary(
            s => s.Key,
            object (s) => s.Value
        );
        MiniExcel.SaveAs(file.FullPath, sheetData);

        return file;
    }

    public static TemporaryFile SaveAsWorkbook(this IEnumerable<Line> lines)
    {
        var file = new TemporaryFile(".xlsx");

        MiniExcel.SaveAs(file.FullPath, lines);

        return file;
    }
}