using System.Diagnostics;
using Microsoft.Win32;

namespace OneNoteSync;

/// <summary>Persisted institution -&gt; OneNote location mapping (onemap.json).</summary>
public sealed class OneMap
{
    public Dictionary<string, MapEntry> Institutions { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>Per-file settings (enabled + title), keyed by file name.</summary>
    public Dictionary<string, FileSetting> Files { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class MapEntry
{
    /// <summary>"section" or "group".</summary>
    public string Kind { get; set; } = "section";
    /// <summary>Section name, or section-group name (a "{group} {year}" section is created in it).</summary>
    public string Name { get; set; } = "";
    /// <summary>The notebook this institution maps to (per-institution, replaces the global notebook).</summary>
    public string Notebook { get; set; } = "";
}

/// <summary>Per-file (per-statement) settings: whether to process it and its page title.</summary>
public sealed class FileSetting
{
    public bool Enabled { get; set; } = true;
    public string Title { get; set; } = "";
}

/// <summary>
/// Shared, UI-agnostic helpers: data location, architecture check, OneNote
/// launch, and the per-statement summary / page-name builders.
/// </summary>
public static class Core
{
    public static string MapPath(string root) => Path.Combine(root, "onemap.json");

    /// <summary>
    /// Find the .env (cwd, then parent) and resolve OUTPUT_DIR relative to it,
    /// so the tool works from the repo root or from a project folder.
    /// </summary>
    public static (string Root, string OutDir, Dictionary<string, string> Env) LocateData()
    {
        // For the GUI, also try the app base directory (so it works when the
        // exe is launched from bin/Debug without a specific cwd).
        string cwd = Directory.GetCurrentDirectory();
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var candidates = new[]
        {
            Path.Combine(cwd, ".env"),
            Path.Combine(cwd, "..", ".env"),
            baseDir,
            Path.Combine(baseDir, ".."),
            Path.Combine(baseDir, "..", ".."),
            Path.Combine(baseDir, "..", "..", ".."),
            Path.Combine(baseDir, "..", "..", "..", ".."),
        }.Select(p => Path.Combine(p, ".env")).Distinct().ToArray();

        string? found = candidates.FirstOrDefault(File.Exists);
        string root = found != null ? Path.GetFullPath(Path.GetDirectoryName(found)!) : cwd;
        var env = EnvFile.Load(found ?? "");
        string outDir = EnvFile.Get(env, "OUTPUT_DIR", "./organized");
        if (!Path.IsPathRooted(outDir)) outDir = Path.GetFullPath(Path.Combine(root, outDir));
        return (root, outDir, env);
    }

    public static void SaveMap(string mapPath, OneMap map)
        => File.WriteAllText(mapPath,
            System.Text.Json.JsonSerializer.Serialize(map,
                new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));

    public static OneMap LoadMap(string mapPath)
        => File.Exists(mapPath)
            ? System.Text.Json.JsonSerializer.Deserialize<OneMap>(File.ReadAllText(mapPath)) ?? new OneMap()
            : new OneMap();

    public static List<StatementFile> LoadStatements(string outDir)
    {
        string statementsPath = Path.Combine(outDir, "statements.json");
        if (!File.Exists(statementsPath))
            throw new Exception($"statements.json not found at {statementsPath}. Run StatementOrganizer first.");
        return System.Text.Json.JsonSerializer
            .Deserialize<List<StatementFile>>(File.ReadAllText(statementsPath))!;
    }

    /// <summary>Build the OneNote page summary as plain text lines.</summary>
    public static List<string> BuildSummaryLines(StatementFile f, Statement st)
    {
        var lines = new List<string>();
        lines.Add($"{(st.StatementDate?.ToString("yyyy-MM") ?? "Statement")} {f.Institution} - {st.AccountName}"
                 + (string.IsNullOrWhiteSpace(st.AccountNumber) ? "" : $" (\u2026{st.AccountNumber})"));
        lines.Add(st.StatementType +
                 (st.StatementDate != null ? $" - issued {st.StatementDate.Value:MMMM d, yyyy}" : ""));
        lines.Add($"Balance: {Money(st.Balance)}");
        if (st.Balance != st.OpeningBalance) lines.Add($"Opening balance: {Money(st.OpeningBalance)}");
        if (st.ClosingBalance.HasValue && (st.Balance == null || st.ClosingBalance.Value != st.Balance))
            lines.Add($"Closing balance: {Money(st.ClosingBalance.Value)}");
        if (st.AsOfDate.HasValue) lines.Add($"As of: {st.AsOfDate.Value:MMMM d, yyyy}");
        if (st.DueDate.HasValue)
            lines.Add($"Payment due: {st.DueDate.Value:MMMM d, yyyy}"
                      + (st.AmountDue.HasValue ? $" - {Money(st.AmountDue.Value)}" : ""));
        foreach (var note in st.Notes) lines.Add("\u2022 " + note);
        return lines;
    }

    public static string Money(decimal? v) => v.HasValue ? v.Value.ToString("N2") : "";

    public static string PageName(Statement st)
    {
        string d = st.StatementDate?.ToString("yyyy-MM") ?? DateTime.Now.ToString("yyyy-MM");
        string n = Sanitize(st.AccountName);
        string name = n.Length > 0 ? $"{d} {n}" : d;
        return name.Length > 60 ? name.Substring(0, 60).Trim() : name;
    }

    /// <summary>
    /// The combined summary for a whole statement file (all its accounts on one
    /// page). Returns the detail lines (the page title is prepended separately).
    /// </summary>
    public static List<string> BuildFileSummary(StatementFile f)
    {
        var lines = new List<string>();
        lines.Add($"{f.Institution} — {f.Statements.Count} account(s)");
        foreach (var st in f.Statements)
        {
            lines.Add("");
            lines.Add($"== {st.AccountName}" + (string.IsNullOrWhiteSpace(st.AccountNumber) ? "" : $" (\u2026{st.AccountNumber})") + " ==");
            lines.Add(st.StatementType +
                     (st.StatementDate != null ? $" — issued {st.StatementDate.Value:MMMM d, yyyy}" : ""));
            lines.Add($"Balance: {Money(st.Balance)}");
            if (st.Balance != st.OpeningBalance) lines.Add($"Opening balance: {Money(st.OpeningBalance)}");
            if (st.ClosingBalance.HasValue && (st.Balance == null || st.ClosingBalance.Value != st.Balance))
                lines.Add($"Closing balance: {Money(st.ClosingBalance.Value)}");
            if (st.AsOfDate.HasValue) lines.Add($"As of: {st.AsOfDate.Value:MMMM d, yyyy}");
            if (st.DueDate.HasValue)
                lines.Add($"Payment due: {st.DueDate.Value:MMMM d, yyyy}"
                          + (st.AmountDue.HasValue ? $" — {Money(st.AmountDue.Value)}" : ""));
            foreach (var note in st.Notes) lines.Add("\u2022 " + note);
        }
        return lines;
    }

    /// <summary>
    /// The page title: "yyyy-MM &lt;title&gt;" (the year-month is prepended, per request).
    /// Uses the latest statement date in the file; falls back to the current month.
    /// </summary>
    public static string PageTitle(StatementFile f, string title)
    {
        var dates = f.Statements.Where(s => s.StatementDate.HasValue).Select(s => s.StatementDate!.Value).ToList();
        string ym = dates.Count > 0 ? dates.Max().ToString("yyyy-MM") : DateTime.Now.ToString("yyyy-MM");
        string t = Sanitize(title);
        if (t.Length == 0) t = Sanitize(f.Institution.Length > 0 ? f.Institution : "Statement");
        string name = $"{ym} {t}";
        return name.Length > 60 ? name.Substring(0, 60).Trim() : name;
    }

    public static string Sanitize(string s)
    {
        var bad = Path.GetInvalidFileNameChars().Concat(new[] { '"', '<', '>', '|' }).ToArray();
        return string.Join(" ", s.Split(bad, StringSplitOptions.RemoveEmptyEntries)).Trim();
    }

    public static string? FindFirstPdf(string outDir)
        => Directory.Exists(outDir)
           ? Directory.EnumerateFiles(outDir, "*.pdf", SearchOption.AllDirectories).FirstOrDefault()
           : null;

    public static int GetDpi(Dictionary<string, string> env)
        => int.TryParse(EnvFile.Get(env, "PDF_DPI", "150"), out int d) ? d : 150;

    /// <summary>Max pages to render as images (PDF_MAX_PAGES, default 5).</summary>
    public static int GetMaxPages(Dictionary<string, string> env)
        => int.TryParse(EnvFile.Get(env, "PDF_MAX_PAGES", "5"), out int m) ? m : 5;

    public static void LaunchOneNote(Action<string>? log = null)
    {
        try
        {
            bool alreadyUp = Process.GetProcessesByName("ONENOTE").Length > 0;
            log?.Invoke($"OneNote already running: {alreadyUp}");
            if (!alreadyUp)
            {
                log?.Invoke("Launching OneNote (onenote.exe) ...");
                Process.Start(new ProcessStartInfo("onenote.exe") { UseShellExecute = true });
            }
            for (int i = 0; i < 30 && Process.GetProcessesByName("ONENOTE").Length == 0; i++)
            {
                log?.Invoke($"  waiting for OneNote process ({i + 1}/30)...");
                Thread.Sleep(1000);
            }
            Thread.Sleep(alreadyUp ? 1000 : 5000);
            log?.Invoke("OneNote process is up; proceeding.");
        }
        catch (Exception e)
        {
            log?.Invoke("LaunchOneNote: " + e.Message + " (COM connect may still start it)");
        }
    }

    /// <summary>
    /// The Office/OneNote COM server only answers a client process of the
    /// same bitness. This app is a .NET Framework executable; by default it
    /// builds AnyCPU which the CLR runs 64-bit on a 64-bit OS. If the
    /// installed Office is 32-bit, every COM call fails with E_FAIL.
    /// </summary>
    public static void CheckArchitecture(Action<string>? log = null)
    {
        bool procIs32 = !Environment.Is64BitProcess;
        bool? officeIs32 = DetectOfficeIs32Bit();
        string procDesc = procIs32 ? "32" : "64";
        string officeDesc = officeIs32 == null ? "unknown" : (officeIs32.Value ? "32" : "64");
        log?.Invoke($"Architecture check: process={procDesc}-bit, office={officeDesc}-bit");
        if (officeIs32 == null || officeIs32 == procIs32) return;

        throw new Exception(
            $"Architecture mismatch: this process is {(procIs32 ? "32" : "64")}-bit, " +
            $"but the installed Office/OneNote is {(officeIs32.Value ? "32" : "64")}-bit. " +
            "The OneNote COM API only works with matching bitness. " +
            (procIs32
                ? "Re-run as 64-bit (this exe is AnyCPU; on a 64-bit OS it is already 64-bit)."
                : "Re-run as 32-bit, or build with /p:PlatformTarget=x86."));
    }

    /// <summary>Detect the bitness of the installed Office/OneNote via the registry.</summary>
    public static bool? DetectOfficeIs32Bit()
    {
        try
        {
            using var software = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            bool hasTop = software.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0") != null;
            bool hasWow = software.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Office\16.0") != null;
            if (Environment.Is64BitProcess)
                return hasTop ? false : (hasWow ? true : null);
            return hasWow ? true : (hasTop ? false : null);
        }
        catch
        {
            return null;
        }
    }
}
