using Shouldly;

namespace Evaluate.ReadingWritingSpreadsheets.Specs.Tooling;

public class WorkbookCreationSpecs
{
    [Fact]
    public void CanCreateWorkbook()
    {
        Line[] lines =
        [
            new() { Name = "Name", Info = "Paul" },
            new() { Name = "City", Info = "Vienna" }
        ];

        // Act
        using var file = lines.SaveAsWorkbook();

        file.HasBeenWritten().ShouldBeTrue();
    }
}