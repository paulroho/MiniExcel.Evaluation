using MiniExcelLibs;

namespace Evaluate.ReadingWritingExcel;

public class ExcelImporter
{
    public List<Line> ImportExcel(string fileName)
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<Line>().ToList();
    }
}