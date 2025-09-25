using MiniExcelLibs;

namespace Evaluate.ReadingWritingExcel;

public class ExcelExporter
{
    public void Write(string fileName, Line[] lines)
    {
        MiniExcel.SaveAs(fileName, lines);
    }
}