using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PaulRoho.Evaluate.ReadingWritingWorkbooks.Specs.Tooling;

public class TemporaryFile : IDisposable
{
    public TemporaryFile(string extension)
    {
        FullPath = $"{Path.GetTempFileName()}{extension}";
        File.Delete(FullPath);
    }

    public bool AutoOpen { get; set; }
    public bool AutoDelete { get; set; } = true;

    public string FullPath { get; }

    public bool HasBeenWritten()
    {
        var fileInfo = new FileInfo(FullPath);
        return fileInfo is
        {
            Exists: true,
            Length: > 0
        };
    }

    public void Dispose()
    {
        if (AutoOpen)
        {
            OpenFile(FullPath);
        }
        else if (AutoDelete)
        {
            File.Delete(FullPath);
        }
    }

    private static void OpenFile(string filePath)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Process.Start(new ProcessStartInfo(filePath)
            {
                UseShellExecute = true
            });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            Process.Start("open", filePath);
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            Process.Start("xdg-open", filePath);
        }
        else
        {
            throw new PlatformNotSupportedException("Unsupported OS");
        }
    }
}