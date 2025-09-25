using Shouldly;

namespace Evaluate.ReadingWritingSpreadsheets.Specs;

public class ReadingSpreadsheetSpecs
{
    [Fact]
    public void CanReadSpreadsheet()
    {
        Line[] data =
        [
            new() { Name = "Name", Info = "It's me" },
            new() { Name = "Height", Info = "Pretty high" }
        ];
        using var file = CreateSpreadsheet(data);
        var importer = new SpreadsheetImporter();

        // Act
        var dataFromFile = importer.ImportSpreadsheet(file.FullPath);

        dataFromFile.ShouldBe(data);
    }

    private TemporaryFile CreateSpreadsheet(IEnumerable<Line> lines)
    {
        var file = new TemporaryFile(".xlsx");
        file.AutoOpen = false;

        var exporter = new SpreadsheetExporter();
        exporter.Write(file.FullPath, lines);

        return file;
    }
}