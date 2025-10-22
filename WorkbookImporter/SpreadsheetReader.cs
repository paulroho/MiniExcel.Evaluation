using System.Linq.Expressions;
using MiniExcelLibs;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks;

public class SpreadsheetReader
{
    public List<T> ReadSpreadsheet<T>(string fileName)
        where T : class, new()
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<T>().ToList();
    }

    public List<T> ReadSpreadsheet<T>(string fileName, string sheetName, string startingCell = "A1")
        where T : class, IRowMarker, new()
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<T>(sheetName, startCell: startingCell)
            .TakeWhile(row => row.IsProcessable)
            .ToList();
    }

    public List<T> ReadSpreadsheet<T>(string fileName, string sheetName, string startingCell,
        Expression<Func<T, string>> mergedValueSelector)
        where T : class, IRowMarker, IWithMergedCell<T>, new()
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<T>(sheetName, startCell: startingCell)
            .TakeWhile(row => row.IsProcessable)
            .UnrollMergedCell(mergedValueSelector)
            .ToList();
    }

    public List<T> ReadSpreadsheet<T>(string fileName, string sheetName, string startingCell,
        Expression<Func<T, string>> mergedValueSelector,
        Expression<Func<T, string>> mergedSubValueSelector)
        where T : class, IRowMarker, IWithMergedCell<T>, IWithMergedSubCell<T>, new()
    {
        using var stream = File.OpenRead(fileName);
        return stream.Query<T>(sheetName, startCell: startingCell)
            .TakeWhile(row => row.IsProcessable)
            .UnrollHierarchicallyMergedCells(mergedValueSelector, mergedSubValueSelector)
            .ToList();
    }


    public interface IRowMarker
    {
        bool IsProcessable { get; }
    }

    public interface IWithMergedCell<out T>
    {
        T WithMergedCellValue(string value);
    }

    public interface IWithMergedSubCell<out T>
    {
        T WithMergedSubCellValue(string value);
    }
}