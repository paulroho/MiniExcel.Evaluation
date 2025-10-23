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

        var previousValue = string.Empty;

        foreach (var row in originalData)
        {
            var currentValue = getMergedValue(row);
            var rowToYield = !string.IsNullOrEmpty(currentValue)
                ? row
                : row.WithMergedCellValue(previousValue);

            previousValue = getMergedValue(rowToYield);

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

        var previousMainValue = string.Empty;
        var previousSubValue = string.Empty;

        foreach (var row in originalData)
        {
            var currentMainValue = getMergedMainValue(row);
            var rowWithMainValue = !string.IsNullOrEmpty(currentMainValue)
                ? row
                : row.WithMergedCellValue(previousMainValue);

            var updatedMainValue = getMergedMainValue(rowWithMainValue);
            if (updatedMainValue != previousMainValue)
            {
                previousSubValue = string.Empty;
            }

            var currentSubValue = getMergedSubValue(rowWithMainValue);
            var rowToYield = !string.IsNullOrEmpty(currentSubValue)
                ? rowWithMainValue
                : rowWithMainValue.WithMergedSubCellValue(previousSubValue);

            previousMainValue = updatedMainValue;
            previousSubValue = getMergedSubValue(rowToYield);

            yield return rowToYield;
        }
    }
}