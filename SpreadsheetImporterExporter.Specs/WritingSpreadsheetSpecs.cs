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
        writer.Write(file.FullPath, lines);

        var fileHasBeenWritten = file.HasBeenWritten();
        fileHasBeenWritten.ShouldBeTrue();
    }
}