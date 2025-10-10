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
        var sheets = new Dictionary<string, object>
        {
            { "Some Sheet", Array.Empty<Line>() },
            { "My Sheet", data },
            { "Another Sheet", Array.Empty<Line>() }
        };
        using var file = CreateSpreadsheet(sheets);

        var reader = new SpreadsheetReader();

        // Act
        var dataFromFile = reader.ReadSpreadsheet(file.FullPath, "My Sheet");

        dataFromFile.ShouldBe(data);
    }

    private static TemporaryFile CreateSpreadsheet(Dictionary<string, object> sheets)
    {
        var file = new TemporaryFile(".xlsx");

        var writer = new SpreadsheetWriter();
        writer.WriteSpreadsheet(file.FullPath, sheets);

        return file;
    }

    private static TemporaryFile CreateSpreadsheet(IEnumerable<Line> lines)
    {
        var file = new TemporaryFile(".xlsx");

        var writer = new SpreadsheetWriter();
        writer.WriteSpreadsheet(file.FullPath, lines);

        return file;
    }
}