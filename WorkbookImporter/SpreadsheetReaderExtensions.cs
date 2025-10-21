using System.Linq.Expressions;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks;

internal static class SpreadsheetReaderExtensions
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

    public static IEnumerable<T> UnrollDoubleMergedCell<T>(
        this IEnumerable<T> originalData,
        Expression<Func<T, string>> mergedCellSelector,
        Expression<Func<T, string>> mergedSubCellSelector)
        where T : SpreadsheetReader.IWithMergedCell<T>, SpreadsheetReader.IWithMergedSubCell<T>
    {
        var getMergedValue = mergedCellSelector.Compile();
        var getMergedSubValue = mergedSubCellSelector.Compile();
        var previousValue = string.Empty;
        var lastValue = string.Empty;
        var lastSubValue = string.Empty;
        foreach (var row in originalData)
        {
            var mergedValue = getMergedValue(row);
            T rowWithProperGroup;
            if (!string.IsNullOrEmpty(mergedValue))
            {
                lastValue = mergedValue;
                rowWithProperGroup = row;
            }
            else
            {
                rowWithProperGroup = row.WithMergedCellValue(lastValue);
            }

            if (getMergedValue(rowWithProperGroup) != previousValue)
            {
                lastSubValue = string.Empty;
            }

            var mergedSubValue = getMergedSubValue(row);
            T rowToReturn;
            if (!string.IsNullOrEmpty(mergedSubValue))
            {
                lastSubValue = mergedSubValue;
                rowToReturn = rowWithProperGroup;
            }
            else
            {
                rowToReturn = rowWithProperGroup.WithMergedSubCellValue(lastSubValue);
            }

            previousValue = lastValue;
            yield return rowToReturn;
        }
    }
}