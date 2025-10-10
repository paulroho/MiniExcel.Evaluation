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
        using var file = data.SaveAsWorkbook();
        var reader = new SpreadsheetReader();

        // Act
        var dataFromFile = reader.ReadSpreadsheet(file.FullPath);

        dataFromFile.ShouldBe(data);
    }

    [Fact]
    public void CanReadSpecificSpreadsheet()
    {
        Line[] data =
        [
            new() { Name = "Name", Info = "It's me" },
            new() { Name = "Height", Info = "Pretty high" }
        ];
        var sheets = new Dictionary<string, Line[]>
        {
            { "Some Sheet", [] },
            { "My Sheet", data },
            { "Another Sheet", [] }
        };
        using var file = sheets.SaveAsWorkbook();

        var reader = new SpreadsheetReader();

        // Act
        var dataFromFile = reader.ReadSpreadsheet(file.FullPath, "My Sheet");

        dataFromFile.ShouldBe(data);
    }
}