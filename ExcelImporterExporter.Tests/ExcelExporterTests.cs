using Shouldly;

namespace Evaluate.ReadingWritingExcel.Tests;

public class ExcelExporterTests
{
    [Fact]
    public void CanCreateExcelFile()
    {
        using var excelFile = new TemporaryFile(".xlsx");
        var exporter = new ExcelExporter();

        Line[] lines =
        [
            new() { Name = "Name", Info = "Paul" },
            new() { Name = "City", Info = "Vienna" }
        ];
        exporter.Write(excelFile.FullPath, lines);

        var fileHasBeenWritten = excelFile.HasBeenWritten();
        fileHasBeenWritten.ShouldBeTrue();
    }
}