using Shouldly;

namespace Evaluate.ReadingWritingExcel.Specs;

public class WritingExcelFileSpecs
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