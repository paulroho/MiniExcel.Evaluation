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

    private record RealisticRow : SpreadsheetReader.IRowMarker
    {
        [ExcelColumnName("Group")] public string Group { get; init; } = string.Empty;
        [ExcelColumnName("")] public string Part { get; init; } = string.Empty;
        [ExcelColumnName("The Text")] public string Text { get; init; } = string.Empty;
        [ExcelColumnName("The Key")] public string Key { get; init; } = string.Empty;
        [ExcelColumnName("The Value")] public decimal Value { get; init; }

        public bool IsProcessable => !string.IsNullOrWhiteSpace(Text);
    }

    [Fact]
    public void CanReadDataWithGrouping()
    {
        // Act
        var dataFromFile = _reader.ReadSpreadsheet<RealisticRow>(
            "SampleFiles/RealisticTemplateFilled.xlsx",
            "Important Sheet",
            "B12");
        dataFromFile = FillFromMergedCell(dataFromFile);

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

    private List<RealisticRow> FillFromMergedCell(List<RealisticRow> originalData)
    {
        var lastGroup = string.Empty;
        var output = new List<RealisticRow>();
        foreach (var row in originalData)
        {
            if (!string.IsNullOrEmpty(row.Group))
            {
                lastGroup = row.Group;
                output.Add(row);
            }
            else
            {
                output.Add(row with { Group = lastGroup });
            }
        }
        return output;
    }
}