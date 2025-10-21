using System.Linq.Expressions;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks;

internal static class SpreadsheetReaderExtensions
{
    public static IEnumerable<T> UnrollMergedCell<T>(
        this IEnumerable<T> originalData,
        Expression<Func<T, string>> mergedValueSelector)
        where T : SpreadsheetReader.IWithMergedCell<T>
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