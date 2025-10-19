using MiniExcelLibs.Attributes;
using PaulRoho.Evaluate.ReadingWritingWorkbooks.Specs.Tooling;
using Shouldly;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks.Specs;

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
        var dataFromFile = reader.ReadSpreadsheet<Line>(file.FullPath, "My Sheet");

        dataFromFile.ShouldBe(data);
    }

    [Theory]
    [InlineData("SampleFiles/TableFromMySheetB3.xlsx")]
    [InlineData("SampleFiles/TableFromMySheetB3_WithFrame.xlsx")]
    public void CanReadSpecificSpreadsheetStartingAtSpecificCell(string fileName)
    {
        Line[] data =
        [
            new() { Name = "Name", Info = "Paul" },
            new() { Name = "Country", Info = "Austria" },
            new() { Name = "City", Info = "Vienna" },
        ];
        var reader = new SpreadsheetReader();

        // Act
        var dataFromFile = reader.ReadSpreadsheet<Line>(fileName, "My Sheet", "B3");

        dataFromFile.ShouldBe(data);
    }

    private record KeyValue
    {
        [ExcelColumnName("The Key")] public string Key { get; init; }
        [ExcelColumnName("The Value")] public decimal Value { get; init; }
    }

    [Fact]
    public void CanReadDecimalValues()
    {
        var reader = new SpreadsheetReader();

        // Act
        var dataFromFile = reader.ReadSpreadsheet<KeyValue>(
            "SampleFiles/TableWithDecimals_ImportantSheet_B10.xlsx", 
            "Important Sheet", 
            "B10");

        dataFromFile.ShouldBe((KeyValue[])[
            new() { Key = "Key A.1", Value = 1.1m },
            new() { Key = "Key A.2", Value = 22.22m },
            new() { Key = "Key A.3", Value = 333.333m },
        ]);
    }
}