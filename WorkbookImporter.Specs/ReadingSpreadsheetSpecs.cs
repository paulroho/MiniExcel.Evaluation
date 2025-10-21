using MiniExcelLibs.Attributes;
using PaulRoho.Evaluate.ReadingWritingWorkbooks.Specs.Tooling;
using Shouldly;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks.Specs;

public class ReadingSpreadsheetSpecs
{
    private readonly SpreadsheetReader _reader;

    public ReadingSpreadsheetSpecs()
    {
        _reader = new SpreadsheetReader();
    }

    // ReSharper disable UnusedAutoPropertyAccessor.Local
    private record Line : SpreadsheetReader.IRowMarker
    {
        public string Name { get; init; } = string.Empty;
        public string Info { get; init; } = string.Empty;
        public bool IsProcessable => !string.IsNullOrWhiteSpace(Name);
    }
    // ReSharper restore UnusedAutoPropertyAccessor.Local

    [Fact]
    public void CanReadSpreadsheet()
    {
        Line[] data =
        [
            new() { Name = "Name", Info = "It's me" },
            new() { Name = "Height", Info = "Pretty high" }
        ];
        using var file = data.SaveAsWorkbook();

        // Act
        var dataFromFile = _reader.ReadSpreadsheet<Line>(file.FullPath);

        dataFromFile.ShouldBe(data);
    }

    [Fact]
    public void CanReadSpecificSpreadsheet()
    {
        object[] data =
        [
            new Line { Name = "Name", Info = "It's me" },
            new Line { Name = "Height", Info = "Pretty high" }
        ];
        var sheets = new Dictionary<string, object[]>
        {
            { "Some Sheet", [] },
            { "My Sheet", data },
            { "Another Sheet", [] }
        };
        using var file = sheets.SaveAsWorkbook();

        // Act
        var dataFromFile = _reader.ReadSpreadsheet<Line>(file.FullPath, "My Sheet");

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

        // Act
        var dataFromFile = _reader.ReadSpreadsheet<Line>(fileName, "My Sheet", "B3");

        dataFromFile.ShouldBe(data);
    }

    // ReSharper disable UnusedAutoPropertyAccessor.Local
    private record KeyValue : SpreadsheetReader.IRowMarker
    {
        [ExcelColumnName("The Key")] public string Key { get; init; } = string.Empty;
        [ExcelColumnName("The Value")] public decimal Value { get; init; }
        public bool IsProcessable => !string.IsNullOrWhiteSpace(Key);
    }
    // ReSharper restore UnusedAutoPropertyAccessor.Local

    [Fact]
    public void CanReadDecimalValues()
    {
        // Act
        var dataFromFile = _reader.ReadSpreadsheet<KeyValue>(
            "SampleFiles/TableWithDecimals_ImportantSheet_B10.xlsx",
            "Important Sheet",
            "B10");

        dataFromFile.ShouldBe((KeyValue[])
        [
            new() { Key = "Key A.1", Value = 1.1m },
            new() { Key = "Key A.2", Value = 22.22m },
            new() { Key = "Key A.3", Value = 333.333m },
        ]);
    }

    [Fact]
    public void CanReadFromFilledSimpleTemplate()
    {
        // Act
        var dataFromFile = _reader.ReadSpreadsheet<KeyValue>(
            "SampleFiles/SimpleTemplateFilled.xlsx",
            "Important Sheet",
            "B10");

        dataFromFile.ShouldBe((KeyValue[])
        [
            new() { Key = "Key A.1", Value = 1.1m },
            new() { Key = "Key A.2", Value = 22.22m },
            new() { Key = "Key A.3", Value = 333.333m },
        ]);
    }
}