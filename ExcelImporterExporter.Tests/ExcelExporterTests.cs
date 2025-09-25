using System.Diagnostics;
using Shouldly;

namespace Evaluate.ReadingWritingExcel.Tests;

public class ExcelExporterTests(ITestOutputHelper testOutputHelper)
{
    [Fact]
    public void CanCreateExcelFile()
    {
        var excelFile = Path.GetTempFileName();
        testOutputHelper.WriteLine($"Created temporary file: {excelFile}");
        var exporter = new ExcelExporter();

        exporter.WriteHello(excelFile);

        var fileHasBeenWritten = FileHasBeenWritten(excelFile);
        fileHasBeenWritten.ShouldBeTrue();

        Open(excelFile);
    }

    private static bool FileHasBeenWritten(string excelFile)
        => new FileInfo(excelFile) is
        {
            Exists: false,
            Length: > 0
        };

    private void Open(string path)
    {
        using var fileopener = new Process();

        fileopener.StartInfo.FileName = "explorer";
        fileopener.StartInfo.Arguments = $"\"{path}\"";
        fileopener.Start();
    }
}