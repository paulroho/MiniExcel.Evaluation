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
            var rowToYield = row.ProcessMergedCell(getMergedValue, lastValue,
                (r, v) => r.WithMergedCellValue(v), out var updatedValue);

            lastValue = updatedValue;

            yield return rowToYield;
        }
    }

    public static IEnumerable<T> UnrollHierarchicallyMergedCells<T>(
        this IEnumerable<T> originalData,
        Expression<Func<T, string>> mergedCellSelector,
        Expression<Func<T, string>> mergedSubCellSelector)
        where T : SpreadsheetReader.IWithMergedCell<T>, SpreadsheetReader.IWithMergedSubCell<T>
    {
        var getMergedMainValue = mergedCellSelector.Compile();
        var getMergedSubValue = mergedSubCellSelector.Compile();

        var lastMainValue = string.Empty;
        var lastSubValue = string.Empty;
        var previousMainValue = string.Empty;

        foreach (var row in originalData)
        {
            var rowWithMainValue = row.ProcessMergedCell(getMergedMainValue, lastMainValue,
                (r, v) => r.WithMergedCellValue(v), out var updatedMainValue);

            if (updatedMainValue != previousMainValue)
            {
                lastSubValue = string.Empty;
            }

            var rowToYield = rowWithMainValue.ProcessMergedCell(getMergedSubValue, lastSubValue,
                (r, v) => r.WithMergedSubCellValue(v), out var updatedSubValue);

            lastMainValue = updatedMainValue;
            lastSubValue = updatedSubValue;
            previousMainValue = updatedMainValue;

            yield return rowToYield;
        }
    }

    private static T ProcessMergedCell<T>(this T row, Func<T, string> getCurrentValue, string fallbackValue,
        Func<T, string, T> withValueFunc, out string updatedValue)
    {
        var currentValue = getCurrentValue(row);

        if (!string.IsNullOrEmpty(currentValue))
        {
            updatedValue = currentValue;
            return row;
        }

        updatedValue = fallbackValue;
        return withValueFunc(row, fallbackValue);
    }
}