using System.Diagnostics;
using System.Text;
using System.Drawing.Imaging;
using PdfiumViewer;

namespace OneNoteSync;

/// <summary>
/// Local PDF rasterization / text extraction.
///
/// Two backends:
///   1. managed PDFium (Apache-2.0 `PdfiumViewer` wrapper + `bblanchon.PDFium`
///      native runtime — open source, ships inside the NuGet packages) —
///      **default**; there is nothing to install.
///   2. poppler-utils (pdftoppm / pdftotext) — opt-in via PDF_BACKEND=poppler
///      (used only when the tools are on PATH).
///
/// The managed backend runs on Windows with zero external dependencies, so the
/// app runs even on a bare Windows machine.
/// </summary>
public static class PdfTools
{
    static PdfTools() => EnsureNativeResolved();

    static void EnsureNativeResolved()
    {
        // The wrapper loads the native pdfium.dll via LoadLibrary; point it at
        // the copy the bblanchon.PDFium package deployed (its own probing does
        // not match the layout the package deploys for .NET Framework, and
        // .NET Framework does not search those subfolders on its own).
        // Architecture-matching copies come first so a wrong-bitness pdfium.dll
        // at the app root is never loaded.
        PdfiumResolver.Resolve += (s, e) =>
        {
            string dir = AppContext.BaseDirectory;
            string preferred = IntPtr.Size == 8 ? "x64" : "x86";
            var candidates = new[]
            {
                Path.Combine(dir, preferred, "pdfium.dll"),
                Path.Combine(dir, "pdfium.dll"),
                Path.Combine(dir, "runtimes", "win-x64", "native", "pdfium.dll"),
                Path.Combine(dir, "runtimes", "win-x86", "native", "pdfium.dll"),
            };
            foreach (string p in candidates)
                if (File.Exists(p)) { e.PdfiumFileName = p; return; }
        };
    }

    /// <summary>Set true to force the poppler backend (env PDF_BACKEND=poppler).
    /// Default is the managed PDFium backend, which is used unless poppler is
    /// requested *and* available.</summary>
    public static bool PreferPoppler = false;

    public static bool HasRasterizer => true;  // managed backend is always available
    public static bool HasTextExtractor => true;

    /// <summary>Rasterizes up to <paramref name="maxPages"/> pages to PNG bytes at <paramref name="dpi"/> DPI.</summary>
    public static List<(string Name, byte[] Png)> Rasterize(string pdfPath, int dpi, int maxPages)
    {
        if (PreferPoppler && HasPoppler("pdftoppm"))
            return RasterizePoppler(pdfPath, dpi, maxPages);
        return RasterizeManaged(pdfPath, dpi, maxPages);
    }

    /// <summary>Extracts the full text of the document (throws if there is no text layer, e.g. scanned).</summary>
    public static string ExtractText(string pdfPath)
    {
        if (PreferPoppler && HasPoppler("pdftotext"))
        {
            var text = RunCapture("pdftotext", $"-layout \"{pdfPath}\" -");
            if (!string.IsNullOrWhiteSpace(text)) return text.Trim();
            // fall through to the managed backend (pdftotext can be empty on weird PDFs)
        }
        return ExtractTextManaged(pdfPath);
    }

    // ---------------------------------------------------------------- poppler

    static List<(string Name, byte[] Png)> RasterizePoppler(string pdfPath, int dpi, int maxPages)
    {
        var work = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "stmt-" + Guid.NewGuid().ToString("N")));
        try
        {
            Run("pdftoppm", $"-png -r {dpi} -f 1 -l {maxPages} \"{pdfPath}\" \"{work.FullName}/page\"");
            var pages = Directory.EnumerateFiles(work.FullName, "*.png").OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
            if (pages.Count == 0) throw new Exception("pdftoppm produced no pages");
            return pages.Select(p => (Path.GetFileNameWithoutExtension(p), File.ReadAllBytes(p))).ToList();
        }
        finally
        {
            try { work.Delete(recursive: true); } catch { /* best effort */ }
        }
    }

    // --------------------------------------------------------- managed (PDFium)

    static List<(string Name, byte[] Png)> RasterizeManaged(string pdfPath, int dpi, int maxPages)
    {
        var pages = new List<(string, byte[])>();
        using var doc = PdfDocument.Load(pdfPath, null);
        int count = Math.Min(doc.PageCount, maxPages);
        for (int i = 0; i < count; i++)
        {
            var sz = doc.PageSizes[i];
            int w = Math.Max(1, (int)(sz.Width * dpi / 72.0));
            int h = Math.Max(1, (int)(sz.Height * dpi / 72.0));

            // The renderer pre-fills the bitmap with white and scales the page
            // by dpi/72; since w/h above are the page size * dpi/72, the page
            // fills the image exactly.
            using var img = doc.Render(i, w, h, dpi, dpi, PdfRotation.Rotate0, PdfRenderFlags.None);
            using var ms = new MemoryStream();
            img.Save(ms, ImageFormat.Png);
            pages.Add(($"page{i + 1}", ms.ToArray()));
        }
        return pages;
    }

    static string ExtractTextManaged(string pdfPath)
    {
        using var doc = PdfDocument.Load(pdfPath, null);
        var sb = new StringBuilder();
        for (int i = 0; i < doc.PageCount; i++)
        {
            sb.Append(doc.GetPdfText(i));
            if (i < doc.PageCount - 1) sb.AppendLine();
        }
        var result = sb.ToString().Trim();
        if (string.IsNullOrWhiteSpace(result))
            throw new Exception("no text layer found (scanned PDF?)");
        return result;
    }

    // ---------------------------------------------------------------- helpers

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

    public static bool HasPoppler(string tool)
    {
        try
        {
            // Windows: `where` finds pdftoppm.exe / pdftotext.exe on PATH.
            // POSIX:   sh -c "command -v ..."
            var psi = new ProcessStartInfo("where", tool) { RedirectStandardOutput = true, UseShellExecute = false };
            using var proc = Process.Start(psi) ?? throw new Exception();
            var outp = proc.StandardOutput.ReadToEnd().Trim();
            proc.WaitForExit(5000);
            return outp.Length > 0;
        }
        catch { return false; }
    }

    static string Truncate(string s, int n) => string.IsNullOrEmpty(s) ? "" : (s.Length <= n ? s : s.Substring(0, n) + "...");
}
