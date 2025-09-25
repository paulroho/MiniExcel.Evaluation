using Shouldly;

namespace Evaluate.ReadingWritingExcel.Tests;

public class ExcelExporterTests
{
    [Fact]
    public void CanCreateExcelFile()
    {
        using var excelFile = new TemporaryFile(".xlsx");
        var exporter = new ExcelExporter();

        exporter.WriteHello(excelFile.FullPath);

        var fileHasBeenWritten = excelFile.HasBeenWritten();
        fileHasBeenWritten.ShouldBeTrue();
    }
}