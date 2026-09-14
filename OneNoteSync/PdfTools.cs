using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using Patagames.Pdf;
using Patagames.Pdf.Net;
using Patagames.Pdf.Enums;

namespace OneNoteSync;

/// <summary>
/// Local PDF rasterization / text extraction.
///
/// Two backends:
///   1. managed Patagames.Pdf (PDFium .NET SDK) — **default**; the native pdfium
///      runtime ships inside the NuGet package (PtgPdfCore.runtime.{linux,osx,win}),
///      so there is nothing to install on any platform.
///   2. poppler-utils (pdftoppm / pdftotext) — opt-in via PDF_BACKEND=poppler
///      (used only when the tools are on PATH).
///
/// The managed backend works on Linux, Windows and macOS with zero external
/// dependencies, so the app runs even on a bare Windows machine.
/// </summary>
public static class PdfTools
{
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
        using var doc = PdfDocument.Load(pdfPath, null, null);
        int count = Math.Min(doc.Pages.Count, maxPages);
        for (int i = 0; i < count; i++)
        {
            var pg = doc.Pages[i];
            var sz = doc.GetPageSizeByIndex(i);
            double scale = dpi / 72.0;
            int w = Math.Max(1, (int)(sz.Width * scale));
            int h = Math.Max(1, (int)(sz.Height * scale));

            // The renderer draws on top of existing pixels, so the bitmap must be
            // pre-filled with white (a fresh bitmap is transparent/black).
            using var bmp = new PdfBitmap(w, h, true);
            bmp.FillRect(0, 0, w, h, FS_COLOR.White);
            pg.Render(bmp, 0, 0, w, h, PageRotate.Normal, RenderFlags.FPDF_NONE);

            int stride = bmp.Stride;
            var raw = new byte[h * stride];
            Marshal.Copy(bmp.Buffer, raw, 0, raw.Length);

            var rows = new MemoryStream();
            for (int y = 0; y < h; y++)
            {
                rows.WriteByte(0); // PNG filter 0 (none)
                rows.Write(raw, y * stride, stride);
            }
            pages.Add(($"page{i + 1}", PngWriter.Encode(rows.ToArray(), w, h)));
        }
        return pages;
    }

    static string ExtractTextManaged(string pdfPath)
    {
        using var doc = PdfDocument.Load(pdfPath, null, null);
        var sb = new StringBuilder();
        for (int i = 0; i < doc.Pages.Count; i++)
        {
            var text = doc.Pages[i].Text.GetText(0, doc.Pages[i].Text.CountChars);
            sb.Append(text);
            if (i < doc.Pages.Count - 1) sb.AppendLine();
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

/// <summary>
/// Minimal PNG encoder (8-bit RGBA, no interlacing, filter 0 on every row).
/// .NET's DeflateStream writes raw DEFLATE, so the zlib wrapper (0x78 0x9C
/// header + Adler-32 trailer) is added manually.
/// <paramref name="rows"/> = per row: 1 filter byte + stride bytes.
/// </summary>
public static class PngWriter
{
    public static byte[] Encode(byte[] rows, int width, int height)
    {
        // zlib-wrapped scanline data
        using var idatMs = new MemoryStream();
        using (var deflate = new DeflateStream(idatMs, CompressionLevel.Optimal, leaveOpen: true))
            deflate.Write(rows, 0, rows.Length);
        var rawDeflate = idatMs.ToArray();
        long adler = Adler32(rows);
        var zlibPayload = new byte[rawDeflate.Length + 6];
        zlibPayload[0] = 0x78; zlibPayload[1] = 0x9C;
        Buffer.BlockCopy(rawDeflate, 0, zlibPayload, 2, rawDeflate.Length);
        zlibPayload[zlibPayload.Length - 4] = (byte)(adler >> 24); zlibPayload[zlibPayload.Length - 3] = (byte)(adler >> 16);
        zlibPayload[zlibPayload.Length - 2] = (byte)(adler >> 8);  zlibPayload[zlibPayload.Length - 1] = (byte)adler;

        using var ms = new MemoryStream();
        void W(byte[] b) => ms.Write(b, 0, b.Length);
        void WBE32(int v) => W(new byte[] { (byte)(v >> 24), (byte)(v >> 16), (byte)(v >> 8), (byte)v });
        void Chunk(string type, byte[] data)
        {
            WBE32(data.Length);
            var tb = Encoding.ASCII.GetBytes(type);
            W(tb); W(data);
            WBE32((int)Crc32(tb.Concat(data).ToArray()));
        }

        W(new byte[] { 0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A });
        var ihdr = new byte[13];
        for (int k = 0; k < 4; k++)
        {
            ihdr[k] = (byte)(width >> (24 - 8 * k));
            ihdr[4 + k] = (byte)(height >> (24 - 8 * k));
        }
        ihdr[8] = 8; // bit depth
        ihdr[9] = 6; // color type: RGBA
        Chunk("IHDR", ihdr);
        Chunk("IDAT", zlibPayload);
        Chunk("IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    static long Adler32(byte[] data)
    {
        uint a = 1, b = 0;
        foreach (var x in data)
        {
            a = (a + x) % 65521;
            b = (b + a) % 65521;
        }
        return (b << 16) | a;
    }

    static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++) crc = (crc >> 1) ^ (0xEDB88320u & (uint)(-(crc & 1)));
        }
        return crc ^ 0xFFFFFFFF;
    }
}
