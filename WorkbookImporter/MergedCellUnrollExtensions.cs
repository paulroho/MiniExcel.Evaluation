using System.Linq.Expressions;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks;

internal static class MergedCellUnrollExtensions
{
    public static IEnumerable<T> UnrollMergedCell<T>(
        this IEnumerable<T> originalData,
        Expression<Func<T, string>> mergedCellSelector)
        where T : SpreadsheetReader.IWithMergedCell<T>
    {
        var getMergedValue = mergedCellSelector.Compile();

        var lastValue = string.Empty;

        foreach (var row in originalData)
        {
            var (rowWithProperGroup, lastVal) = row.WithValueForMergedCell(getMergedValue, lastValue);

            lastValue = lastVal;

            yield return rowWithProperGroup;
        }
    }

    public static IEnumerable<T> UnrollHierarchiclyMergedCells<T>(
        this IEnumerable<T> originalData,
        Expression<Func<T, string>> mergedCellSelector,
        Expression<Func<T, string>> mergedSubCellSelector)
        where T : SpreadsheetReader.IWithMergedCell<T>, SpreadsheetReader.IWithMergedSubCell<T>
    {
        var getMergedValue = mergedCellSelector.Compile();
        var getMergedSubValue = mergedSubCellSelector.Compile();

        var lastValue = string.Empty;
        var lastSubValue = string.Empty;
        var previousValue = string.Empty;

        foreach (var row in originalData)
        {
            var (rowWithProperGroup, lastVal) = row.WithValueForMergedCell(getMergedValue, lastValue);

            if (getMergedValue(rowWithProperGroup) != previousValue)
            {
                lastSubValue = string.Empty;
            }

            var (rowToReturn, lastSubVal) =
                rowWithProperGroup.WithValueForMergedSubCell(getMergedSubValue, lastSubValue);

            lastValue = lastVal;
            lastSubValue = lastSubVal;
            previousValue = lastValue;

            yield return rowToReturn;
        }
    }

    private static (T, string) WithValueForMergedCell<T>(this T row, Func<T, string> getMergedValue, string lastValue)
        where T : SpreadsheetReader.IWithMergedCell<T>
    {
        var mergedValue = getMergedValue(row);
        return !string.IsNullOrEmpty(mergedValue)
            ? (row, mergedValue)
            : (row.WithMergedCellValue(lastValue), lastValue);
    }

    private static (T, string) WithValueForMergedSubCell<T>(this T row, Func<T, string> getMergedValue,
        string lastValue)
        where T : SpreadsheetReader.IWithMergedSubCell<T>
    {
        var mergedValue = getMergedValue(row);
        return !string.IsNullOrEmpty(mergedValue)
            ? (row, mergedValue)
            : (row.WithMergedSubCellValue(lastValue), lastValue);
    }
}