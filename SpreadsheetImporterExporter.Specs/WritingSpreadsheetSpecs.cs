using Shouldly;

namespace Evaluate.ReadingWritingSpreadsheets.Specs;

public class WritingSpreadsheetSpecs
{
    [Fact]
    public void CanCreateSpreadsheet()
    {
        using var file = new TemporaryFile(".xlsx");
        var writer = new SpreadsheetWriter();

        Line[] lines =
        [
            new() { Name = "Name", Info = "Paul" },
            new() { Name = "City", Info = "Vienna" }
        ];
        
        // Act
        writer.WriteSpreadsheet(file.FullPath, lines);

        file.HasBeenWritten().ShouldBeTrue();
    }
}