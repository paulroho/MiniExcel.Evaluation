using System.Linq.Expressions;
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
    private record RealisticRow : SpreadsheetReader.IRowMarker, IWithMergedCell<RealisticRow>
    {
        [ExcelColumnName("Group")] public string Group { get; init; } = string.Empty;
        [ExcelColumnName("")] public string Part { get; init; } = string.Empty;
        [ExcelColumnName("The Text")] public string Text { get; init; } = string.Empty;
        [ExcelColumnName("The Key")] public string Key { get; init; } = string.Empty;
        [ExcelColumnName("The Value")] public decimal Value { get; init; }

        public bool IsProcessable => !string.IsNullOrWhiteSpace(Text);
        public RealisticRow WithMergedCellValue(string value) => this with { Group = value };
    }
    // ReSharper restore UnusedAutoPropertyAccessor.Local

    [Fact]
    public void CanReadDataWithGrouping()
    {
        // Act
        var dataFromFile = _reader.ReadSpreadsheet<RealisticRow>(
            "SampleFiles/RealisticTemplateFilled.xlsx",
            "Important Sheet",
            "B12");
        dataFromFile = FillFromMergedCell(dataFromFile, g => g.Group).ToList();

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

    private static IEnumerable<T> FillFromMergedCell<T>(
        IEnumerable<T> originalData,
        Expression<Func<T, string>> mergedValueSelector)
        where T : IWithMergedCell<T>
    {
        var getMergedValue = mergedValueSelector.Compile();
        var lastValue = string.Empty;
        foreach (var row in originalData)
        {
            var mergedValue = getMergedValue(row);
            if (!string.IsNullOrEmpty(mergedValue))
            {
                lastValue = mergedValue;
                yield return row;
            }
            else
            {
                yield return row.WithMergedCellValue(lastValue);
            }
        }
    }
}

internal interface IWithMergedCell<T>
{
    T WithMergedCellValue(string value);
}