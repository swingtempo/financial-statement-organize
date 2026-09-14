using System.Diagnostics;
using System.Net;
using System.Text;
using Microsoft.Win32;
using StatementOrganizer;

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
            if (OperatingSystem.IsWindows() == false)
            {
                Console.Error.WriteLine("OneNoteSync is Windows-only (OneNote desktop COM API).");
                return 1;
            }

            if (args.Length == 0 || args[0] == "--list" || args[0] == "--test")
                CheckArchitecture();

            if (args.Length > 0 && args[0] == "--list")
                return ListSections();
            if (args.Length > 0 && args[0] == "--test")
                return Test(outDir, env, args);
            if (args.Length > 0 && args[0] == "--dry-run")
                return Run(root, outDir, env, dryRun: true);
            return Run(root, outDir, env, dryRun: false);
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("Error: " + e.Message);
            if (e is System.Runtime.InteropServices.COMException ce)
                Console.Error.WriteLine($"HResult: 0x{ce.HResult:X8}");
            return 1;
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

    /// <summary>
    /// Prototype for the risky part: page creation, embedded printout
    /// images, summary text, and PDF attachment. Run this first to verify
    /// visually in OneNote before a full run.
    /// </summary>
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
        Console.WriteLine($"Creating test page '{pageName}' in [{sec.Notebook}] {sec.Name} ...");
        string uri = onenote.CreateNote(sec.Guid, pageName);

        // Sample summary block.
        var sb = new StringBuilder();
        sb.Append("<html><body>");
        sb.Append("<h1>OneNoteSync test</h1>");
        sb.Append("<h2>Example Account — ****1234</h2>");
        sb.Append("<p>Sample statement — issued June 30, 2026</p>");
        sb.Append("<table border=\"1\" cellspacing=\"0\">");
        sb.Append("<tr><td>Balance</td><td>$26,372.47</td></tr>");
        sb.Append("<tr><td>Opening balance</td><td>$23,056.08</td></tr>");
        sb.Append("</table>");

        // Generated test image (gradient rectangle).
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
        sb.Append("<p>Rasterized image test (should show a color gradient):</p>");
        sb.Append($"<img src=\"data:image/png;base64,{Convert.ToBase64String(png)}\" width=\"600\"/>");

        // Real PDF printout (first page) if a PDF is available.
        string? pdf = FindFirstPdf(outDir);
        if (pdf != null)
        {
            Console.WriteLine($"Rasterizing {Path.GetFileName(pdf)} for the printout test ...");
            var pages = PdfTools.Rasterize(pdf, GetDpi(env), 1);
            if (pages.Count > 0)
            {
                sb.Append("<p>PDF page 1 printout:</p>");
                sb.Append($"<img src=\"data:image/png;base64,{Convert.ToBase64String(pages[0].Png)}\" width=\"800\"/>");
            }
        }
        sb.Append("</body></html>");

        onenote.UpdatePageHtml(uri, sb.ToString());

        if (pdf != null)
        {
            onenote.AddFilesToPage(uri, pdf);
            Console.WriteLine($"Attached {Path.GetFileName(pdf)}");
        }
        else
        {
            Console.WriteLine("No PDF found under the output dir — skipping attachment test.");
        }

        Console.WriteLine();
        Console.WriteLine("Done. Check OneNote: page '" + pageName + "' in section '" + sec.Name + "'.");
        Console.WriteLine("You should see: the sample summary, the color gradient image," +
                        (pdf != null ? " the PDF page printout, and the PDF attached to the page." : "."));
        Console.WriteLine("If all of that looks right, run without --test for the full import.");
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

        int dpi = GetDpi(env);
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

            var pages = PdfTools.Rasterize(pdf, dpi, MaxPrintoutPages);
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
                string uri = onenote.CreateNote(sec.Guid, name);
                onenote.UpdatePageHtml(uri, BuildPageHtml(f, st, pages));
                onenote.AddFilesToPage(uri, pdf);
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
    /// same bitness. .NET runs 64-bit by default; if the installed Office
    /// is 32-bit, every COM call fails with E_FAIL (0x80004005).
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
            ? "Fix: dotnet publish -c Release -r win-x64 --self-contained, " +
              "then run bin\\Release\\net9.0\\win-x64\\publish\\OneNoteSync.exe"
            : "Fix: dotnet publish -c Release -r win-x86 --self-contained, " +
              "then run bin\\Release\\net9.0\\win-x86\\publish\\OneNoteSync.exe");
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

            onenote.CreateSection(name);
            h = onenote.GetHierarchy();
            var created = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                         ?? throw new Exception($"Failed to create section \"{name}\".");
            onenote.MoveSectionToGroup(created.Guid, g.Guid);
            h = onenote.GetHierarchy();
            return h.Sections.First(s => s.Guid == created.Guid);
        }

        SectionInfo? sec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
        if (sec != null) return sec;

        onenote.CreateSection(entry.Name);
        h = onenote.GetHierarchy();
        return h.Sections.First(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Build the OneNote page HTML: summary above the PDF printout.</summary>
    private static string BuildPageHtml(StatementFile f, Statement st, List<(string Name, byte[] Png)> pages)
    {
        var sb = new StringBuilder();
        sb.Append("<html><body>");
        sb.Append("<h1>").Append(E((st.StatementDate?.ToString("yyyy-MM") ?? "Statement")))
          .Append(" — ").Append(E(f.Institution)).Append("</h1>");
        sb.Append("<h2>").Append(E(st.AccountName));
        if (!string.IsNullOrWhiteSpace(st.AccountNumber))
            sb.Append(" (…").Append(E(st.AccountNumber)).Append(")");
        sb.Append("</h2>");
        sb.Append("<p>").Append(E(st.StatementType));
        if (st.StatementDate != null) sb.Append(" — issued ").Append(E(st.StatementDate.Value.ToString("MMMM d, yyyy")));
        sb.Append("</p>");

        sb.Append("<table border=\"1\" cellspacing=\"0\">");
        AppendMoneyRow(sb, "Balance", st.Balance);
        if (st.Balance != st.OpeningBalance) AppendMoneyRow(sb, "Opening balance", st.OpeningBalance);
        if (st.ClosingBalance != null && (st.Balance == null || st.ClosingBalance != st.Balance))
            AppendMoneyRow(sb, "Closing balance", st.ClosingBalance);
        if (st.AsOfDate != null) AppendRow(sb, "As of", E(st.AsOfDate.Value.ToString("MMMM d, yyyy")));
        if (st.DueDate != null)
            AppendRow(sb, "Payment due", E(st.DueDate.Value.ToString("MMMM d, yyyy")) +
                     (st.AmountDue.HasValue ? " — " + Money(st.AmountDue.Value) : ""));
        sb.Append("</table>");

        if (st.Notes.Count > 0)
        {
            sb.Append("<ul>");
            foreach (var note in st.Notes) sb.Append("<li>").Append(E(note)).Append("</li>");
            sb.Append("</ul>");
        }

        if (pages.Count > 0)
        {
            sb.Append("<h3>Statement printout</h3>");
            for (int i = 0; i < pages.Count; i++)
            {
                sb.Append($"<p>Page {i + 1} of {pages.Count}</p>");
                sb.Append($"<img src=\"data:image/png;base64,{Convert.ToBase64String(pages[i].Png)}\" width=\"800\"/>");
            }
        }
        sb.Append("</body></html>");
        return sb.ToString();
    }

    private static void AppendRow(StringBuilder sb, string label, string value)
        => sb.Append($"<tr><td>{label}</td><td>{value}</td></tr>");

    private static void AppendMoneyRow(StringBuilder sb, string label, decimal? value)
    {
        if (value.HasValue) AppendRow(sb, label, Money(value.Value));
    }

    private static string Money(decimal v) => v.ToString("N2");

    private static string E(string s) => WebUtility.HtmlEncode(s ?? "");

    private static string PageName(Statement st)
    {
        string d = st.StatementDate?.ToString("yyyy-MM") ?? DateTime.Now.ToString("yyyy-MM");
        string n = Sanitize(st.AccountName);
        string name = n.Length > 0 ? $"{d} {n}" : d;
        return name.Length > 60 ? name[..60].Trim() : name;
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
            Process.Start(new ProcessStartInfo("onenote.exe") { UseShellExecute = true });
            // Give OneNote a moment to register its COM server; the COM
            // connect below would auto-start it anyway, but a fresh launch
            // can race otherwise.
            Thread.Sleep(3000);
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
