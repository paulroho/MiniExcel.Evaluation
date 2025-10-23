using MiniExcelLibs.Attributes;
using Shouldly;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks.Specs;

public class ReadingRealisticSpreadsheetSpecs
{
    private readonly SpreadsheetReader _reader;

    public ReadingRealisticSpreadsheetSpecs()
    {
        _reader = new SpreadsheetReader();
    }

    // ReSharper disable UnusedAutoPropertyAccessor.Local
    private record RealisticRow : SpreadsheetReader.IRowMarker,
        SpreadsheetReader.IWithMergedCell<RealisticRow>,
        SpreadsheetReader.IWithMergedSubCell<RealisticRow>
    {
        [ExcelColumnName("Group")] public string Group { get; init; } = string.Empty;
        [ExcelColumnName("")] public string Part { get; init; } = string.Empty;
        [ExcelColumnName("The Text")] public string Text { get; init; } = string.Empty;
        [ExcelColumnIndex(4)] public string Key { get; init; } = string.Empty;
        [ExcelColumnName("The Value")] public decimal Value { get; init; }

        public bool IsProcessable => !string.IsNullOrWhiteSpace(Text);
        public RealisticRow WithMergedCellValue(string value) => this with { Group = value };
        public RealisticRow WithMergedSubCellValue(string value) => this with { Part = value };
    }
    // ReSharper restore UnusedAutoPropertyAccessor.Local

    [Fact]
    public void CanReadDataWithGrouping()
    {
        // Act
        var dataFromFile = _reader.ReadSpreadsheet<RealisticRow>(
            "SampleFiles/RealisticTemplateFilled.xlsx",
            "Important Sheet",
            "B12",
            row => row.Group);

        dataFromFile.Count.ShouldBe(3);

        dataFromFile[0].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group A.a",
            Part = "",
            Text = "Text A.1",
            Key = "A_a_A1",
            Value = 1.1m
        });
        dataFromFile[1].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group A.a",
            Part = "",
            Text = "Text A.2",
            Key = "A_a_A2",
            Value = 2.22m
        });
        dataFromFile[2].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group A.b",
            Part = "",
            Text = "Text A.3",
            Key = "A_b_A3",
            Value = 3.333m
        });
    }

    [Fact]
    public void CanReadDataWithGroupingAndPart()
    {
        // Act
        var dataFromFile = _reader.ReadSpreadsheet<RealisticRow>(
            "SampleFiles/RealisticTemplateFilled.xlsx",
            "Important Sheet",
            "B18",
            row => row.Group,
            row => row.Part);

        dataFromFile.Count.ShouldBe(5);

        dataFromFile[0].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group B.a",
            Part = "Part 1",
            Text = "Text B.1",
            Key = "B_a_P1_B1",
            Value = 555.55m
        });
        dataFromFile[1].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group B.a",
            Part = "Part 1",
            Text = "Text B.2",
            Key = "B_a_P1_B2",
            Value = 44.44m
        });
        dataFromFile[2].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group B.a",
            Part = "Part 2",
            Text = "Text B.3",
            Key = "B_a_P2_B3",
            Value = 3.33m
        });
        dataFromFile[3].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group B.a",
            Part = "Part 2",
            Text = "Text B.4",
            Key = "B_a_P2_B4",
            Value = 2.2m
        });
        dataFromFile[4].ShouldBeEquivalentTo(new RealisticRow
        {
            Group = "Group B.b",
            Part = "",
            Text = "Text B.5",
            Key = "B_b_B5",
            Value = 1m
        });
    }
}