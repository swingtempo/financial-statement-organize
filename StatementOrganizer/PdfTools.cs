using System.Diagnostics;

namespace StatementOrganizer;

/// <summary>
/// Local PDF rasterization / text extraction.
///
/// Current implementation shells out to poppler-utils (pdftoppm / pdftotext),
/// which is what runs on Linux. poppler has no official Windows build, so for
/// Windows this class is the single seam to swap implementations behind the
/// same interface (e.g. Quest.Pdf or Aspose.Pdf when available, or a poppler
/// build for Windows / WSL).
/// </summary>
public static class PdfTools
{
    public static bool HasRasterizer => ToolExists("pdftoppm");
    public static bool HasTextExtractor => ToolExists("pdftotext");

    /// <summary>Rasterizes up to <paramref name="maxPages"/> pages to PNG bytes at <paramref name="dpi"/> DPI.</summary>
    public static List<(string Name, byte[] Png)> Rasterize(string pdfPath, int dpi, int maxPages)
    {
        var work = Directory.CreateTempSubdirectory("stmt-");
        try
        {
            var prefix = work.FullName + "/page";
            Run("pdftoppm", $"-png -r {dpi} -f 1 -l {maxPages} \"{pdfPath}\" \"{prefix}\"");

            var pages = Directory.EnumerateFiles(work.FullName, "*.png").OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
            if (pages.Count == 0) throw new Exception("pdftoppm produced no pages");
            return pages.Select(p => (Path.GetFileNameWithoutExtension(p), File.ReadAllBytes(p))).ToList();
        }
        finally
        {
            try { work.Delete(recursive: true); } catch { /* best effort */ }
        }
    }

    /// <summary>Extracts the full text of the document ("" if the PDF has no text layer, e.g. scanned).</summary>
    public static string ExtractText(string pdfPath)
    {
        var text = RunCapture("pdftotext", $"-layout \"{pdfPath}\" -");
        if (string.IsNullOrWhiteSpace(text))
            throw new Exception("pdftotext extracted no text (scanned PDF?)");
        return text.Trim();
    }

    static void Run(string tool, string arguments)
    {
        var psi = new ProcessStartInfo(tool, arguments) { RedirectStandardError = true };
        using var proc = Process.Start(psi) ?? throw new Exception($"{tool} failed to start");
        var err = proc.StandardError.ReadToEnd();
        proc.WaitForExit();
        if (proc.ExitCode != 0) throw new Exception($"{tool} failed ({proc.ExitCode}): {Truncate(err, 200)}");
    }

    static string RunCapture(string tool, string arguments)
    {
        var psi = new ProcessStartInfo(tool, arguments)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        using var proc = Process.Start(psi) ?? throw new Exception($"{tool} failed to start");
        var output = proc.StandardOutput.ReadToEnd();
        var err = proc.StandardError.ReadToEnd();
        proc.WaitForExit();
        if (proc.ExitCode != 0 && string.IsNullOrWhiteSpace(output))
            throw new Exception($"{tool} failed ({proc.ExitCode}): {Truncate(err, 200)}");
        return output;
    }

    static bool ToolExists(string tool)
    {
        try
        {
            var psi = new ProcessStartInfo("sh", new[] { "-c", $"command -v {tool}" });
            psi.RedirectStandardOutput = true;
            using var proc = Process.Start(psi) ?? throw new Exception();
            var outp = proc.StandardOutput.ReadToEnd().Trim();
            proc.WaitForExit(5000);
            return outp.Length > 0;
        }
        catch { return false; }
    }

    static string Truncate(string s, int n) => string.IsNullOrEmpty(s) ? "" : (s.Length <= n ? s : s[..n] + "...");
}
