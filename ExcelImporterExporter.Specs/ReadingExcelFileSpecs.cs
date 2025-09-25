using Shouldly;

namespace Evaluate.ReadingWritingExcel.Specs;

public class ReadingExcelFileSpecs
{
    [Fact]
    public void CanReadExcelFile()
    {
        Line[] data =
        [
            new() { Name = "Name", Info = "It's me" },
            new() { Name = "Height", Info = "Pretty high" }
        ];
        using var excelFile = CreateExcelFile(data);
        var importer = new ExcelImporter();

        // Act
        var dataFromFile = importer.ImportExcel(excelFile.FullPath);

        dataFromFile.ShouldBe(data);
    }

    private TemporaryFile CreateExcelFile(IEnumerable<Line> lines)
    {
        var excelFileName = new TemporaryFile(".xlsx");
        excelFileName.AutoOpen = false;

        var exporter = new ExcelExporter();
        exporter.Write(excelFileName.FullPath, lines);

        return excelFileName;
    }
}