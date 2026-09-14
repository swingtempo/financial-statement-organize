using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OneNoteSync;

/// <summary>
/// Windows Forms front-end for OneNoteSync. Shows the institution -&gt;
/// OneNote location mapping (editable), the available sections / section
/// groups, and a live debug log of everything the sync does.
/// </summary>
public sealed class MainForm : Form
{
    private readonly string[] _args;

    private string _root = "";
    private string _outDir = "";
    private Dictionary<string, string> _env = new();
    private List<StatementFile> _files = new();
    private OneMap _map = new();
    private Hierarchy? _hierarchy;

    private bool _busy;
    private bool _updating;

    // Controls.
    private Panel _top = null!;
    private Button _btnReload = null!;
    private Button _btnDryRun = null!;
    private Button _btnTest = null!;
    private Button _btnRun = null!;
    private Button _btnSave = null!;
    private Button _btnClear = null!;
    private Label _info = null!;
    private DataGridView _grid = null!;
    private DataGridViewComboBoxColumn _colTarget = null!;
    private DataGridViewTextBoxColumn _colInst = null!;
    private DataGridViewTextBoxColumn _colKind = null!;
    private DataGridViewTextBoxColumn _colStatus = null!;
    private RichTextBox _log = null!;
    private readonly List<string> _pendingLog = new();

    public MainForm(string[] args)
    {
        _args = args ?? Array.Empty<string>();
        BuildUi();
        // Route any Console.* output (e.g. OneNote.cs internal debug writes)
        // into the GUI log so nothing is lost in the WinExe (no console window).
        var cw = new ConsoleToLog(this);
        Console.SetOut(cw);
        Console.SetError(cw);
        LoadData();
        Shown += (_, _) => FlushPendingLog();
    }

    // ------------------------------------------------------------------
    // UI construction (no designer, kept self-contained)
    // ------------------------------------------------------------------

    private void BuildUi()
    {
        Text = "OneNoteSync — financial statements → OneNote";
        Width = 1000;
        Height = 720;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9f);

        // Top action bar.
        _top = new Panel { Dock = DockStyle.Top, Height = 48 };
        var x = 10;
        _btnReload = AddButton(_top, "Reload OneNote", ref x);
        _btnReload.Click += (_, _) => RunSta(DoReloadOneNote);
        _btnDryRun = AddButton(_top, "Dry run", ref x);
        _btnDryRun.Click += (_, _) => RunSta(() => DoRun(dryRun: true));
        _btnTest = AddButton(_top, "Test page", ref x);
        _btnTest.Click += (_, _) => RunSta(DoTest);
        _btnRun = AddButton(_top, "Run sync", ref x);
        _btnRun.Click += (_, _) => RunSta(() => DoRun(dryRun: false));
        _btnSave = AddButton(_top, "Save mapping", ref x);
        _btnSave.Click += (_, _) => SaveMappingFromGrid();
        _btnClear = AddButton(_top, "Clear log", ref x);
        _btnClear.Click += (_, _) => { if (_log.IsHandleCreated) _log.Clear(); };

        // Info line.
        _info = new Label { Dock = DockStyle.Top, Height = 24, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0) };

        // Log (bottom).
        _log = new RichTextBox
        {
            Dock = DockStyle.Bottom,
            Height = 200,
            ReadOnly = true,
            Multiline = true,
            WordWrap = false,
            ScrollBars = RichTextBoxScrollBars.ForcedBoth,
            BackColor = Color.White,
            Font = new Font("Consolas", 9f),
        };

        // Mapping grid (fills the rest).
        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = SystemColors.Window,
        };
        _colInst = new DataGridViewTextBoxColumn { HeaderText = "Institution", ReadOnly = true, FillWeight = 22 };
        _colTarget = new DataGridViewComboBoxColumn { HeaderText = "Target (section or section-group)", FillWeight = 40 };
        _colKind = new DataGridViewTextBoxColumn { HeaderText = "Kind", ReadOnly = true, FillWeight = 12 };
        _colStatus = new DataGridViewTextBoxColumn { HeaderText = "Status", ReadOnly = true, FillWeight = 12 };
        _grid.Columns.AddRange(_colInst, _colTarget, _colKind, _colStatus);
        _grid.CellValueChanged += Grid_CellValueChanged;

        // Order matters for Dock layout: add Fill last, then bottom/top.
        Controls.Add(_grid);
        Controls.Add(_log);
        Controls.Add(_info);
        Controls.Add(_top);
    }

    private static Button AddButton(Panel parent, string text, ref int x)
    {
        var b = new Button { Text = text, Location = new Point(x, 8), Width = 110, Height = 32, AutoSize = false };
        parent.Controls.Add(b);
        x += 118;
        return b;
    }

    // ------------------------------------------------------------------
    // Initial data load
    // ------------------------------------------------------------------

    private void LoadData()
    {
        try
        {
            var (root, outDir, env) = Core.LocateData();
            _root = root; _outDir = outDir; _env = env;
            Log($"Data root: {root}");
            Log($"Output dir: {outDir}");
            _files = Core.LoadStatements(outDir);
            Log($"Loaded {_files.Count} statement file(s) from statements.json.");

            string mapPath = Core.MapPath(root);
            _map = Core.LoadMap(mapPath);
            Log($"Mapping file: {mapPath} ({_map.Institutions.Count} institution(s) mapped)");

            RefreshInfo();
            RefreshGrid();
        }
        catch (Exception e)
        {
            Log("Startup error: " + e.Message, error: true);
            if (e.InnerException != null) Log("  inner: " + e.InnerException.Message, error: true);
        }
    }

    private void RefreshInfo()
    {
        int goodFiles = _files.Count(f => !f.Error && f.Statements.Count > 0);
        int stmts = _files.Sum(f => f.Statements.Count);
        _info.Text = $"  {_files.Count} files / {stmts} statements ({goodFiles} usable)   •   " +
                     $"OneNote: {(_hierarchy == null ? "not loaded (click Reload OneNote)" : $"{_hierarchy.Sections.Count} sections, {_hierarchy.Groups.Count} groups")}";
    }

    private void RefreshGrid()
    {
        _updating = true;
        // Rebuild the target options from the loaded hierarchy.
        _colTarget.Items.Clear();
        var options = new List<string>();
        if (_hierarchy != null)
        {
            foreach (var g in _hierarchy.Groups) options.Add(g.Name);
            foreach (var s in _hierarchy.Sections) options.Add(s.Name);
        }
        _colTarget.Items.AddRange(options.ToArray());

        _grid.Rows.Clear();
        var institutions = _files
            .Where(f => !f.Error && f.Statements.Count > 0)
            .Select(f => string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase);

        foreach (var inst in institutions)
        {
            string target = "";
            if (_map.Institutions.TryGetValue(inst, out var e)) target = e.Name;
            int r = _grid.Rows.Add(inst, target, DeriveKind(target),
                                   string.IsNullOrWhiteSpace(target) ? "UNMAPPED" : "mapped");
            // Highlight unmapped rows.
            _grid.Rows[r].DefaultCellStyle.BackColor = string.IsNullOrWhiteSpace(target)
                ? Color.LightYellow : Color.White;
        }
        _updating = false;
    }

    private string DeriveKind(string target)
    {
        if (string.IsNullOrWhiteSpace(target)) return "";
        string t = target.Trim();
        if (_hierarchy != null && _hierarchy.Groups.Any(g => g.Name.Equals(t, StringComparison.OrdinalIgnoreCase)))
            return "group";
        return "section";
    }

    private void Grid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (_updating || e.RowIndex < 0) return;
        if (e.ColumnIndex == _colTarget.Index)
        {
            string target = _grid.Rows[e.RowIndex].Cells[_colTarget.Index].Value?.ToString() ?? "";
            _updating = true;
            _grid.Rows[e.RowIndex].Cells[_colKind.Index].Value = DeriveKind(target);
            bool mapped = !string.IsNullOrWhiteSpace(target);
            _grid.Rows[e.RowIndex].Cells[_colStatus.Index].Value = mapped ? "mapped" : "UNMAPPED";
            _grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = mapped ? Color.White : Color.LightYellow;
            _updating = false;
        }
    }

    /// <summary>Read the mapping for one institution from the grid.</summary>
    private (bool Ok, MapEntry Entry) GetMappingForGrid(string inst)
    {
        foreach (DataGridViewRow row in _grid.Rows)
        {
            if (row.Cells[_colInst.Index].Value?.ToString() == inst)
            {
                string target = row.Cells[_colTarget.Index].Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(target)) return (false, new MapEntry());
                string kind = (row.Cells[_colKind.Index].Value?.ToString() ?? DeriveKind(target)).ToLowerInvariant();
                if (kind != "group") kind = "section";
                return (true, new MapEntry { Kind = kind, Name = target.Trim() });
            }
        }
        return (false, new MapEntry());
    }

    // ------------------------------------------------------------------
    // Background execution (STA thread, since OneNote COM is STA)
    // ------------------------------------------------------------------

    private void RunSta(Action action)
    {
        if (_busy) { Log("Already busy — waiting for the current operation to finish.", error: true); return; }
        SetBusy(true);
        var t = new Thread(() =>
        {
            try { action(); }
            catch (Exception e)
            {
                Log("Operation failed: " + e.GetType().Name + ": " + e.Message, error: true);
                if (e is System.Runtime.InteropServices.COMException ce)
                    Log("  HRESULT 0x" + ce.HResult.ToString("X8"), error: true);
                if (e.InnerException != null) Log("  inner: " + e.InnerException.Message, error: true);
            }
            finally { SetBusy(false); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
    }

    private void SetBusy(bool busy)
    {
        try
        {
            if (_log == null || !_log.IsHandleCreated) { _busy = busy; return; }
            if (_log.InvokeRequired) { _log.BeginInvoke(new Action(() => SetBusy(busy))); return; }
            _busy = busy;
            foreach (var b in new[] { _btnReload, _btnDryRun, _btnTest, _btnRun, _btnSave })
                b.Enabled = !busy;
            _info.Text = (busy ? "  WORKING…  " : "") + _info.Text;
        }
        catch { /* form already closing */ }
    }

    private void EnsureHierarchy(OneNote onenote)
    {
        if (_hierarchy == null)
        {
            Log("OneNote hierarchy not loaded — loading now.");
            _hierarchy = onenote.GetHierarchy();
        }
    }

    // ------------------------------------------------------------------
    // Actions
    // ------------------------------------------------------------------

    private void DoReloadOneNote()
    {
        Core.CheckArchitecture(s => Log(s));
        Core.LaunchOneNote(s => Log(s));
        using var onenote = new OneNote();
        Log("Reading OneNote hierarchy (this can take a few seconds) ...");
        var h = onenote.GetHierarchy();
        _hierarchy = h;
        Log($"Loaded {h.Sections.Count} section(s), {h.Groups.Count} section group(s).");
        RefreshGrid();
        RefreshInfo();
    }

    private void DoRun(bool dryRun)
    {
        if (_files.Count == 0)
        {
            Log("No statement files loaded — run StatementOrganizer first.", error: true);
            return;
        }

        OneNote? onenote = null;
        Hierarchy? h = null;
        if (!dryRun)
        {
            Core.CheckArchitecture(s => Log(s));
            Core.LaunchOneNote(s => Log(s));
            onenote = new OneNote();
            h = onenote.GetHierarchy();
            _hierarchy = h;
        }

        int made = 0, skipped = 0, noTarget = 0;
        Log($"Importing {_files.Count} file(s) {(dryRun ? "(dry run)" : "into OneNote")}");

        foreach (var f in _files.Where(f => !f.Error && f.Statements.Count > 0))
        {
            string inst = string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution;
            var (ok, entry) = GetMappingForGrid(inst);
            if (!ok)
            {
                Log($"  {inst}: NO TARGET SET in the mapping — skipping {f.Statements.Count} statement(s).", error: true);
                skipped += f.Statements.Count;
                noTarget++;
                continue;
            }

            string pdf = Path.Combine(_outDir, f.Category, f.FileName);
            if (!File.Exists(pdf))
            {
                Log($"  ! {f.FileName}: PDF not found at {pdf} — skipping.", error: true);
                skipped += f.Statements.Count;
                continue;
            }

            foreach (var st in f.Statements)
            {
                int year = st.StatementDate?.Year ?? DateTime.Now.Year;
                string pageName = Core.PageName(st);

                if (dryRun)
                {
                    string label = entry.Kind == "group" ? $"{entry.Name} {year}" : entry.Name;
                    Log($"  [plan] {f.FileName} :: {st.AccountName} -> {label} / {pageName}");
                    continue;
                }

                Log($"  {f.FileName} :: {st.AccountName}");
                var sec = ResolveSection(onenote!, ref h!, entry, year);
                Log($"    -> [{sec.Notebook}] {sec.Name} / {pageName}");
                string pageId = onenote!.CreateNote(sec.Guid);
                var lines = Core.BuildSummaryLines(f, st);
                lines.Add("PDF: " + pdf);
                onenote.CommitPage(pageId, pageName, lines);
                made++;
            }
        }

        onenote?.Dispose();

        if (dryRun)
        {
            Log($"Dry run complete. {noTarget} institution(s) had no target set.");
        }
        else
        {
            Log($"Done. {made} OneNote page(s) created, {skipped} statement(s) skipped.");
            SaveMappingFromGrid(silent: true);
            Log("OneNote is open so you can browse the new pages.");
        }
        RefreshInfo();
    }

    private void DoTest()
    {
        Core.CheckArchitecture(s => Log(s));
        Core.LaunchOneNote(s => Log(s));
        using var onenote = new OneNote();
        EnsureHierarchy(onenote);
        if (_hierarchy!.Sections.Count == 0)
        {
            Log("No OneNote sections found. Create a section in OneNote first, then click Reload OneNote.", error: true);
            return;
        }

        var sec = _hierarchy.Sections[0];
        string pageName = "OneNoteSync test " + DateTime.Now.ToString("yy-MM-dd HH:mm");
        Log($"Creating test page in [{sec.Notebook}] {sec.Name} ...");
        string pageId = onenote.CreateNote(sec.Guid);

        var lines = new List<string>
        {
            "OneNoteSync test - sample account",
            "Account: Example Checking (\u20261234)",
            "Statement type: Monthly statement - issued June 30, 2026",
            "Balance: $26,372.47",
            "Opening balance: $23,056.08",
        };

        // Generated gradient image so the printout path is exercised even without a PDF.
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
        string imgB64 = Convert.ToBase64String(PngWriter.Encode(rows, w, ht));

        string? pdf = Core.FindFirstPdf(_outDir);
        string? fileB64 = null;
        if (pdf != null)
        {
            Log("Using real PDF for the test: " + pdf);
            lines.Add("PDF: " + pdf);
            fileB64 = Convert.ToBase64String(File.ReadAllBytes(pdf));
            var pages = PdfTools.Rasterize(pdf, Core.GetDpi(_env), 1);
            if (pages.Count > 0) imgB64 = Convert.ToBase64String(pages[0].Png);
        }

        Log("Committing (core always; binary is best-effort) ...");
        var (imgOk, fileOk, note) = onenote.TryCommitPageWithBinary(
            pageId, pageName, lines, imgB64, w, ht, fileB64,
            pdf != null ? Path.GetFileName(pdf) : null);

        Log($"Test done: page '{pageName}' in section '{sec.Name}'.");
        Log("  Core (title + summary text): always committed.");
        Log("  " + note + ".");
        if (imgOk || fileOk)
            Log("  A binary part registered - verify the image/attachment render in OneNote.");
        else
            Log("  Binary part not registered on this machine (known 2013-schema limitation). " +
                "Title, summary, and the PDF path are saved; open the PDF from the path line.");
    }

    /// <summary>
    /// Resolve (and create, if needed) the target section for a mapping entry.
    /// For a group mapping, the section is "{group} {year}".
    /// </summary>
    private static SectionInfo ResolveSection(OneNote onenote, ref Hierarchy h, MapEntry entry, int year)
    {
        if (entry.Kind == "group")
        {
            var g = h.Groups.FirstOrDefault(g => g.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase))
                  ?? throw new Exception($"Section group \"{entry.Name}\" no longer exists in OneNote.");
            string name = $"{entry.Name} {year}";

            var existing = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (existing != null) return existing;

            onenote.CreateSectionInGroup(name, g.Guid);
            h = onenote.GetHierarchy();
            return h.Sections.First(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        var sec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
        if (sec != null) return sec;

        onenote.CreateSection(entry.Name);
        h = onenote.GetHierarchy();
        return h.Sections.First(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
    }

    private void SaveMappingFromGrid(bool silent = false)
    {
        var map = new OneMap();
        foreach (DataGridViewRow row in _grid.Rows)
        {
            string inst = row.Cells[_colInst.Index].Value?.ToString() ?? "";
            string target = row.Cells[_colTarget.Index].Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(target)) continue; // skip unmapped
            string kind = (row.Cells[_colKind.Index].Value?.ToString() ?? DeriveKind(target)).ToLowerInvariant();
            if (kind != "group") kind = "section";
            map.Institutions[inst] = new MapEntry { Kind = kind, Name = target.Trim() };
        }
        try
        {
            string mapPath = Core.MapPath(_root);
            Core.SaveMap(mapPath, map);
            if (!silent) Log($"Saved {map.Institutions.Count} mapping(s) to {mapPath}");
        }
        catch (Exception e)
        {
            Log("Save mapping failed: " + e.Message, error: true);
        }
    }

    // ------------------------------------------------------------------
    // Logging
    // ------------------------------------------------------------------

    public void Log(string msg, bool error = false)
    {
        string line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        try
        {
            if (_log == null || !_log.IsHandleCreated)
            {
                _pendingLog.Add(line + "|" + (error ? "e" : ""));
                return;
            }
            if (_log.InvokeRequired)
                _log.BeginInvoke(new Action(() => AppendLog(line, error)));
            else
                AppendLog(line, error);
        }
        catch { /* form already closing — drop the message rather than crash */ }
    }

    private void AppendLog(string line, bool error)
    {
        _log.SelectionStart = _log.TextLength;
        _log.SelectionColor = error ? Color.Firebrick : Color.Black;
        _log.AppendText(line + Environment.NewLine);
        _log.SelectionColor = Color.Black;
        _log.SelectionStart = _log.TextLength;
        _log.SelectionLength = 0;
        _log.ScrollToCaret();
    }

    private void FlushPendingLog()
    {
        foreach (var entry in _pendingLog)
        {
            int bar = entry.LastIndexOf('|');
            string line = entry.Substring(0, bar);
            bool error = entry.Substring(bar + 1) == "e";
            AppendLog(line, error);
        }
        _pendingLog.Clear();
    }

    /// <summary>Writes Console.Out / Console.Error into the GUI log (thread-safe).</summary>
    private sealed class ConsoleToLog : TextWriter
    {
        private readonly MainForm _form;
        public ConsoleToLog(MainForm form) { _form = form; }
        public override Encoding Encoding => Encoding.UTF8;
        public override void Write(char ch) { /* partial; the line arrives via Write(string) */ }
        public override void Write(string? value)
        {
            if (string.IsNullOrEmpty(value)) return;
            foreach (var piece in value.Split('\n'))
            {
                string line = piece.TrimEnd('\r');
                if (line.Length > 0) _form.Log("  | " + line);
            }
        }
    }
}
