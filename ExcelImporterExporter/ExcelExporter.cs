using MiniExcelLibs;

namespace Evaluate.ReadingWritingExcel;

public class ExcelExporter
{
    public void WriteHello(string fileName)
    {
        MiniExcel.SaveAs(fileName,
            new[]
            {
                new { Name = "Name", Info = "Paul" },
                new { Name = "City", Info = "Vienna" }
            });
    }
}