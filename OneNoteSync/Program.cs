using System.Diagnostics;
using System.Net;
using System.Text;
using Microsoft.Win32;

namespace OneNoteSync;

public static class Program
{
    /// <summary>
    /// Persists institution -> OneNote location so the user is only asked
    /// once per new institution (across runs).
    /// </summary>
    private sealed class OneMap
    {
        public Dictionary<string, MapEntry> Institutions { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private sealed class MapEntry
    {
        /// <summary>"section" or "group".</summary>
        public string Kind { get; set; } = "section";
        public string Name { get; set; } = "";
    }

    private const int MaxPrintoutPages = 20;

    public static int Main(string[] args)
    {
        try
        {
            var (root, outDir, env) = LocateData();
            if (!System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
            {
                Console.Error.WriteLine("OneNoteSync is Windows-only (OneNote desktop COM API).");
                return 1;
            }

            if (args.Length == 0 || args[0] == "--list" || args[0] == "--test")
                CheckArchitecture();

            if (args.Length > 0 && args[0] == "--list")
                return ListSections();
            if (args.Length > 0 && args[0] == "--diag")
                return Diag();
            if (args.Length > 0 && args[0] == "--test")
                return Test(outDir, env, args);
            if (args.Length > 0 && args[0] == "--dry-run")
                return Run(root, outDir, env, dryRun: true);
            return Run(root, outDir, env, dryRun: false);
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("Error: " + e.Message);
            Console.Error.WriteLine(OneNote.DescribeHr(e));
            return 1;
        }
    }

    // ------------------------------------------------------------------
    // Diagnostics
    // ------------------------------------------------------------------

    private static int Diag()
    {
        Console.WriteLine("=== OneNote COM diagnostics ===");
        Console.WriteLine($"Process bitness: {(Environment.Is64BitProcess ? "64" : "32")}-bit");

        // 1. What does the ProgID resolve to, and what type does it produce?
        try
        {
            var t = Type.GetTypeFromProgID("OneNote.Application");
            Console.WriteLine($"Type.GetTypeFromProgID -> {t}");
            Console.WriteLine($"  Declaring assembly: {t.Assembly.FullName}");
        }
        catch (Exception e) { Console.WriteLine("  GetTypeFromProgID FAILED: " + e.Message); }

        // 2. Verify the PIA object answers GetHierarchy.
        try
        {
            using var onenote = new OneNote();
            var raw = onenote.GetRawHierarchy();
            Console.WriteLine($"[OK] PIA GetHierarchy raw length = {raw.Length}");
            var head = raw.Length > 2500 ? raw.Substring(0, 2500) : raw;
            Console.WriteLine("---- raw XML head ----");
            Console.WriteLine(head);
            Console.WriteLine("----------------------");
            var h = onenote.GetHierarchy();
            Console.WriteLine($"[OK] parsed -> {h.Sections.Count} sections, {h.Groups.Count} section groups");
        }
        catch (Exception e)
        {
            DumpChain(e, "PIA GetHierarchy");
        }
        return 0;

    }

    private static void Probe(string label, Func<string> call)
    {
        try
        {
            string? r = call();
            Console.WriteLine($"[OK]   {label} -> {(r ?? "")}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"[FAIL] {label}: {e.GetType().Name}: {e.Message}");
            var inner = e.InnerException;
            if (inner != null) Console.WriteLine($"       inner: {inner.GetType().Name}: {inner.Message}");
        }
    }

    private static void TryCall(Type t, string label, object[] args)
    {
        try
        {
            t.InvokeMember("GetHierarchy".Equals("x") ? "GetHierarchy" : MethodFromLabel(label), System.Reflection.BindingFlags.InvokeMethod, null, null, args);
            string? outval = args.LastOrDefault() as string;
            Console.WriteLine($"[OK]   {label} -> out={(outval == null ? "(none)" : outval.Length + " chars")}");
        }
        catch (System.Reflection.TargetInvocationException tie)
        {
            var inner = tie.InnerException ?? tie;
            Console.WriteLine($"[FAIL] {label}: {inner.GetType().Name}: {inner.Message}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"[FAIL] {label}: {e.GetType().Name}: {e.Message}");
        }
    }
    private static string MethodFromLabel(string label) => label.Split('(')[0];

    private static void DumpChain(Exception e, string label)
    {
        Console.WriteLine($"--- {label}: {e.GetType().Name}: {e.Message}");
        var hr = e is System.Runtime.InteropServices.COMException ce ? ce.HResult
                 : e is System.Runtime.InteropServices.SEHException se ? se.HResult : 0;
        Console.WriteLine($"  HResult: {(hr == 0 ? "(none)" : $"0x{hr:X8}")}");
        var inner = e.InnerException;
        int depth = 0;
        while (inner != null && depth++ < 5)
        {
            Console.WriteLine($"  inner[{depth}]: {inner.GetType().Name}: {inner.Message}");
            if (inner is System.Runtime.InteropServices.COMException ic) Console.WriteLine($"    inner HResult: 0x{ic.HResult:X8}");
            inner = inner.InnerException;
        }
    }

    // ------------------------------------------------------------------
    // CLI commands
    // ------------------------------------------------------------------

    private static int ListSections()
    {
        LaunchOneNote();
        using var onenote = new OneNote();
        var h = onenote.GetHierarchy();
        Console.WriteLine("Section groups:");
        foreach (var g in h.Groups)
            Console.WriteLine($"  [GROUP] [{g.Notebook}] {g.Name}");
        Console.WriteLine();
        Console.WriteLine("Sections:");
        foreach (var s in h.Sections)
            Console.WriteLine($"  [{s.Notebook}] {s.Name}");
        Console.WriteLine($"\n({h.Groups.Count} groups, {h.Sections.Count} sections)");
        return 0;
    }

    
    
    

    private static int Test(string outDir, Dictionary<string, string> env, string[] args)
    {
        LaunchOneNote();
        using var onenote = new OneNote();
        var h = onenote.GetHierarchy();
        if (h.Sections.Count == 0)
        {
            Console.Error.WriteLine("No OneNote sections found. Create a section in OneNote first.");
            return 1;
        }

        string? want = args.Length > 1 ? args[1] : null;
        SectionInfo sec = want == null
            ? h.Sections[0]
            : h.Sections.FirstOrDefault(s => s.Name.Equals(want, StringComparison.OrdinalIgnoreCase))
              ?? throw new Exception($"Section \"{want}\" not found. Use --list to see names.");

        string pageName = "OneNoteSync test " + DateTime.Now.ToString("yy-MM-dd HH:mm");
        Console.WriteLine($"Creating test page in [{sec.Notebook}] {sec.Name} ...");
        string pageId = onenote.CreateNote(sec.Guid);

        // Sample summary lines (the always-working core).
        var lines = new List<string>
        {
            "OneNoteSync test - sample account",
            "Account: Example Checking (\u20261234)",
            "Statement type: Monthly statement - issued June 30, 2026",
            "Balance: $26,372.47",
            "Opening balance: $23,056.08",
        };

        // Generated test image (gradient rectangle) so the printout path is
        // exercised even without a PDF.
        int w = 600, ht = 200;
        var rows = new byte[w * ht * 3];
        for (int y = 0; y < ht; y++)
            for (int x = 0; x < w; x++)
            {
                int i = (y * w + x) * 3;
                rows[i] = (byte)(x * 255 / w);
                rows[i + 1] = (byte)(y * 255 / ht);
                rows[i + 2] = 180;
            }
        byte[] png = PngWriter.Encode(rows, w, ht);
        string imgB64 = Convert.ToBase64String(png);

        // Real PDF printout + attachment, if a PDF is available.
        string? pdf = FindFirstPdf(outDir);
        string? fileB64 = null;
        if (pdf != null)
        {
            lines.Add("PDF: " + pdf);
            fileB64 = Convert.ToBase64String(File.ReadAllBytes(pdf));
            var pages = PdfTools.Rasterize(pdf, GetDpi(env), 1);
            if (pages.Count > 0) imgB64 = Convert.ToBase64String(pages[0].Png); // use real PDF page 1
        }

        Console.WriteLine("Committing (core always; binary is the prototype) ...");
        var (imgOk, fileOk, note) = onenote.TryCommitPageWithBinary(
            pageId, pageName, lines, imgB64, w, ht, fileB64,
            pdf != null ? Path.GetFileName(pdf) : null);

        Console.WriteLine();
        Console.WriteLine("Done. Check OneNote: page " + pageName + " in section " + sec.Name + ".");
        Console.WriteLine("  Core (title + summary text): always committed.");
        Console.WriteLine("  " + note + ".");
        if (imgOk || fileOk)
        {
            Console.WriteLine("  A binary part registered - verify the image/attachment render in OneNote.");
            Console.WriteLine("  If they look right, run without --test for the full import.");
        }
        else
        {
            Console.WriteLine("  The OneNote COM API could not register the binary part on this");
            Console.WriteLine("  machine (a known limitation of the 2013 page schema). The title,");
            Console.WriteLine("  summary, and the PDF path are saved; open the PDF from the path line.");
        }
        return 0;
    }

    // ------------------------------------------------------------------
    // Full run
    // ------------------------------------------------------------------

    private static int Run(string root, string outDir, Dictionary<string, string> env, bool dryRun)
    {
        string statementsPath = Path.Combine(outDir, "statements.json");
        if (!File.Exists(statementsPath))
            throw new Exception($"statements.json not found at {statementsPath}. Run StatementOrganizer first.");

        var files = System.Text.Json.JsonSerializer
            .Deserialize<List<StatementFile>>(File.ReadAllText(statementsPath))!;

        string mapPath = Path.Combine(root, "onemap.json");
        var map = File.Exists(mapPath)
            ? System.Text.Json.JsonSerializer.Deserialize<OneMap>(File.ReadAllText(mapPath))!
            : new OneMap();

        int made = 0, skipped = 0;

        if (!dryRun) LaunchOneNote();
        using var onenote = dryRun ? null : new OneNote();
        var h = dryRun
            ? new Hierarchy(Array.Empty<SectionInfo>(), Array.Empty<GroupInfo>())
            : onenote!.GetHierarchy();

        Console.WriteLine($"Importing {files.Count} files from {outDir} " + (dryRun ? "(dry run)" : "into OneNote"));
        Console.WriteLine();

        foreach (var f in files.Where(f => !f.Error && f.Statements.Count > 0))
        {
            string inst = string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution;
            MapEntry entry;
            if (!map.Institutions.TryGetValue(inst, out entry))
            {
                if (dryRun)
                {
                    entry = new MapEntry { Kind = "section", Name = "(prompt)" };
                    Console.WriteLine($"  {inst}: NEW institution (would prompt for section/group)");
                }
                else
                {
                    entry = PromptForInstitution(inst, onenote!, ref h);
                    map.Institutions[inst] = entry;
                }
            }

            string pdf = Path.Combine(outDir, f.Category, f.FileName);
            if (!File.Exists(pdf))
            {
                Console.WriteLine($"  ! {f.FileName}: PDF not found at {pdf} — skipped");
                skipped += f.Statements.Count;
                continue;
            }

            foreach (var st in f.Statements)
            {
                int year = st.StatementDate?.Year ?? DateTime.Now.Year;

                if (dryRun)
                {
                    string label = entry.Kind == "group" ? $"{entry.Name} {year}" : entry.Name;
                    Console.WriteLine($"  [plan] {f.FileName} :: {st.AccountName} -> {label} / {PageName(st)}");
                    continue;
                }

                Console.WriteLine($"  {f.FileName} :: {st.AccountName} (pending section resolve)");
                SectionInfo sec = ResolveSection(onenote!, ref h, entry, year);
                string name = PageName(st);
                Console.WriteLine($"    -> [{sec.Notebook}] {sec.Name} / {name}");
                string pageId = onenote.CreateNote(sec.Guid);
                var lines = BuildSummaryLines(f, st);
                lines.Add("PDF: " + pdf);
                onenote.CommitPage(pageId, name, lines);
                made++;
            }
        }

        if (!dryRun && map.Institutions.Count > 0)
            SaveMap(mapPath, map);

        Console.WriteLine();
        Console.WriteLine(dryRun
            ? "Dry run complete."
            : $"Done. {made} OneNote page(s) created, {skipped} statement(s) skipped. OneNote is open so you can browse.");
        return 0;
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    /// <summary>
    /// The Office/OneNote COM server only answers a client process of the
    /// same bitness. This app is a .NET Framework executable; by default
    /// it builds AnyCPU which the CLR runs 64-bit on a 64-bit OS. If the
    /// installed Office is 32-bit, every COM call fails with E_FAIL.
    /// Detect the mismatch up front and tell the user how to re-run.
    /// </summary>
    private static void CheckArchitecture()
    {
        bool procIs32 = !Environment.Is64BitProcess;
        bool? officeIs32 = DetectOfficeIs32Bit();
        if (officeIs32 == null || officeIs32 == procIs32) return;

        Console.Error.WriteLine($"Architecture mismatch: this process is {(procIs32 ? "32" : "64")}-bit, " +
                              $"but the installed Office/OneNote is {(officeIs32.Value ? "32" : "64")}-bit. " +
                              "The OneNote COM API only works with matching bitness.");
        Console.Error.WriteLine(procIs32
            ? "Fix: re-run as 64-bit (this exe is AnyCPU; on a 64-bit OS it is already 64-bit)."
            : "Fix: re-run as 32-bit, e.g. run the 32-bit OneNoteSync.exe, or build with /p:PlatformTarget=x86.");
        throw new Exception("Office/OneNote architecture mismatch (see above). No changes were made.");
    }

    /// <summary>Detect the bitness of the installed Office/OneNote via the registry.</summary>
    private static bool? DetectOfficeIs32Bit()
    {
        try
        {
            // From a 64-bit view of HKLM: top-level Office\16.0 = 64-bit Office,
            // WOW6432Node\Office\16.0 = 32-bit Office.
            using var software = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            bool hasTop = software.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0") != null;
            bool hasWow = software.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Office\16.0") != null;
            // If both bitnesses are installed, report the one matching this process
            // (the COM server of the matching bitness is the one we will use).
            if (Environment.Is64BitProcess)
                return hasTop ? false : (hasWow ? true : null);
            return hasWow ? true : (hasTop ? false : null);
        }
        catch
        {
            return null; // unknown — let the COM call speak
        }
    }

    /// <summary>
    /// Ask the user which section or section group an institution's
    /// statements should go into.
    /// </summary>
    private static MapEntry PromptForInstitution(string inst, OneNote onenote, ref Hierarchy h)
    {
        if (Console.IsInputRedirected)
        {
            Console.Error.WriteLine("No interactive input available; cannot prompt for a new institution.");
            throw new Exception("Run interactively to map institution '" + inst + "'.");
        }

        var sections = h.Sections.ToList();
        var groups = h.Groups.ToList();

        while (true)
        {
            Console.WriteLine();
            Console.WriteLine($"Which OneNote location for \"{inst}\"?");
            Console.WriteLine();
            Console.WriteLine("Sections:");
            for (int i = 0; i < sections.Count; i++)
                Console.WriteLine($"  {i + 1,2}) [{sections[i].Notebook}] {sections[i].Name}");
            Console.WriteLine("Section groups (a section named \"{group} {year}\" will be created in the group):");
            for (int i = 0; i < groups.Count; i++)
                Console.WriteLine($"  {sections.Count + i + 1,2}) [GROUP] [{groups[i].Notebook}] {groups[i].Name}");
            Console.WriteLine("     n) New section (you choose the name)");

            Console.Write("> ");
            string? line = Console.ReadLine()?.Trim();
            if (line == "n")
            {
                Console.Write("New section name: ");
                string name = Console.ReadLine()?.Trim() ?? "";
                if (name.Length == 0) continue;
                return new MapEntry { Kind = "section", Name = name };
            }

            if (int.TryParse(line, out int n) && n >= 1)
            {
                if (n <= sections.Count)
                    return new MapEntry { Kind = "section", Name = sections[n - 1].Name };
                if (n <= sections.Count + groups.Count)
                    return new MapEntry { Kind = "group", Name = groups[n - 1 - sections.Count].Name };
            }
            Console.WriteLine("Invalid choice — try again.");
        }
    }

    /// <summary>
    /// Resolve (and create, if needed) the target section for an institution
    /// mapping entry. For a group mapping, the section is "{group} {year}".
    /// </summary>
    private static SectionInfo ResolveSection(OneNote onenote, ref Hierarchy h, MapEntry entry, int year)
    {
        if (entry.Kind == "group")
        {
            var g = h.Groups.FirstOrDefault(g => g.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase))
                  ?? throw new Exception($"Section group \"{entry.Name}\" no longer exists in OneNote.");
            string name = $"{entry.Name} {year}";

            SectionInfo? existing = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (existing != null) return existing;

            // Create the section directly inside the section group (no move needed).
            onenote.CreateSectionInGroup(name, g.Guid);
            h = onenote.GetHierarchy();
            return h.Sections.First(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        SectionInfo? sec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
        if (sec != null) return sec;

        onenote.CreateSection(entry.Name);
        h = onenote.GetHierarchy();
        return h.Sections.First(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Build the OneNote page summary as plain text lines.</summary>
    private static List<string> BuildSummaryLines(StatementFile f, Statement st)
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

    private static string Money(decimal? v) => v.HasValue ? v.Value.ToString("N2") : "";

    private static string PageName(Statement st)
    {
        string d = st.StatementDate?.ToString("yyyy-MM") ?? DateTime.Now.ToString("yyyy-MM");
        string n = Sanitize(st.AccountName);
        string name = n.Length > 0 ? $"{d} {n}" : d;
        return name.Length > 60 ? name.Substring(0, 60).Trim() : name;
    }

    private static string Sanitize(string s)
    {
        var bad = Path.GetInvalidFileNameChars().Concat(new[] { '"', '<', '>', '|' }).ToArray();
        return string.Join(" ", s.Split(bad, StringSplitOptions.RemoveEmptyEntries)).Trim();
    }

    private static void SaveMap(string mapPath, OneMap map)
        => File.WriteAllText(mapPath,
            System.Text.Json.JsonSerializer.Serialize(map,
                new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));

    private static string? FindFirstPdf(string outDir)
        => Directory.Exists(outDir)
           ? Directory.EnumerateFiles(outDir, "*.pdf", SearchOption.AllDirectories).FirstOrDefault()
           : null;

    private static int GetDpi(Dictionary<string, string> env)
        => int.TryParse(EnvFile.Get(env, "PDF_DPI", "150"), out int d) ? d : 150;

    private static void LaunchOneNote()
    {
        try
        {
            bool alreadyUp = Process.GetProcessesByName("ONENOTE").Length > 0;
            if (!alreadyUp)
                Process.Start(new ProcessStartInfo("onenote.exe") { UseShellExecute = true });
            // Wait for the OneNote process to be up and give the COM server a
            // moment to finish initializing (its first calls often fail otherwise).
            for (int i = 0; i < 30 && Process.GetProcessesByName("ONENOTE").Length == 0; i++)
                Thread.Sleep(1000);
            Thread.Sleep(alreadyUp ? 1000 : 5000);
        }
        catch
        {
            // COM connect will start it anyway.
        }
    }

    /// <summary>
    /// Find the .env (cwd, then parent) and resolve OUTPUT_DIR relative to it,
    /// so the tool works from the repo root or from a project folder.
    /// </summary>
    private static (string Root, string OutDir, Dictionary<string, string> Env) LocateData()
    {
        string cwd = Directory.GetCurrentDirectory();
        string? found = new[] { Path.Combine(cwd, ".env"), Path.Combine(cwd, "..", ".env") }
            .FirstOrDefault(File.Exists);
        string root = found != null ? Path.GetFullPath(Path.GetDirectoryName(found)!) : cwd;
        var env = EnvFile.Load(found ?? "");
        string outDir = EnvFile.Get(env, "OUTPUT_DIR", "./organized");
        if (!Path.IsPathRooted(outDir)) outDir = Path.GetFullPath(Path.Combine(root, outDir));
        return (root, outDir, env);
    }
}
