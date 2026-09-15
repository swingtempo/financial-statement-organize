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
    private System.Threading.CancellationTokenSource? _cts;

    // Controls.
    private Panel _top = null!;
    private Button _btnReload = null!;
    private Button _btnDryRun = null!;
    private Button _btnTest = null!;
    private Button _btnRun = null!;
    private Button _btnStop = null!;
    private Button _btnToggleAll = null!;
    private Button _btnSave = null!;
    private Button _btnClear = null!;
    private Label _info = null!;
    private DataGridView _grid = null!;
    private DataGridViewCheckBoxColumn _colEnabled = null!;
    private DataGridViewTextBoxColumn _colInst = null!;
    private DataGridViewTextBoxColumn _colFile = null!;
    private DataGridViewComboBoxColumn _colNotebook = null!;
    private DataGridViewComboBoxColumn _colTarget = null!;
    private DataGridViewButtonColumn _colOpen = null!;
    private DataGridViewTextBoxColumn _colTitle = null!;
    private DataGridViewTextBoxColumn _colKind = null!;
    private DataGridViewTextBoxColumn _colStatus = null!;
    private RichTextBox _log = null!;
    private readonly List<string> _pendingLog = new();

    /// <summary>
    /// A pickable target. Carries the real section/section-group name, its kind
    /// (so it can be indicated in the dropdown), and its owning notebook (for
    /// the notebook filter). Displayed as "[GROUP] name" / "[SECTION] name".
    /// </summary>
    private sealed class TargetOption
    {
        public string Name { get; }
        public string Kind { get; }        // "section" or "group"
        public string Notebook { get; }
        public string Parent { get; }      // containing section-group (for sections; the notebook for groups)
        public TargetOption(string name, string kind, string notebook, string parent = "")
        {
            Name = name; Kind = kind; Notebook = notebook; Parent = parent;
        }
        public override string ToString()
        {
            string nb = string.IsNullOrEmpty(Notebook) ? "" : $"  ({Notebook})";
            return Kind == "group" ? $"[GROUP] {Name}{nb}" : $"[SECTION] {Name}{nb}";
        }
        // Value equality (Name + Kind + Notebook + Parent). The combo cell holds an
        // instance built by one list and the Items hold instances built by another;
        // reference equality would make the grid report "value is not valid".
        // Including Notebook + Parent disambiguates same-name sections in different
        // groups / notebooks.
        public override bool Equals(object? obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (obj is not TargetOption other) return false;
            return string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Kind, other.Kind, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Notebook ?? "", other.Notebook ?? "", StringComparison.OrdinalIgnoreCase)
                && string.Equals(Parent ?? "", other.Parent ?? "", StringComparison.OrdinalIgnoreCase);
        }
        public override int GetHashCode()
        {
            unchecked
            {
                int h = (Name ?? "").GetHashCode();
                h = (h * 397) ^ (Kind ?? "").GetHashCode();
                h = (h * 397) ^ (Notebook ?? "").GetHashCode();
                h = (h * 397) ^ (Parent ?? "").GetHashCode();
                return h;
            }
        }
    }

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
        // Auto-load OneNote at boot (per request): launch OneNote if needed, connect,
        // load the hierarchy, and refresh the grid — no manual "Reload OneNote" click.
        // Runs on a background STA thread so it doesn't block the UI; the connect-retry
        // in OneNoteApp rides out a slow OneNote start.
        _ = System.Threading.Tasks.Task.Run(() =>
        {
            RunSta(() =>
            {
                try
                {
                    Log("[boot] Auto-loading OneNote ...");
                    DoReloadOneNote();
                    Log("[boot] OneNote auto-load done.");
                }
                catch (Exception e) { Log("[boot] OneNote auto-load issue: " + e.Message); }
            });
        });
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

        // Top action bar (single row: buttons).
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
        _btnStop = AddButton(_top, "Stop", ref x);
        _btnStop.Click += (_, _) => { _cts?.Cancel(); Log("Stop requested — finishing the current page, then stopping."); };
        _btnStop.Enabled = false;
        _btnToggleAll = AddButton(_top, "Toggle all", ref x);
        _btnToggleAll.Click += (_, _) => ToggleAll();
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
        _colEnabled = new DataGridViewCheckBoxColumn { HeaderText = "On", FillWeight = 6 };
        _colInst = new DataGridViewTextBoxColumn { HeaderText = "Institution", ReadOnly = true, FillWeight = 18 };
        _colFile = new DataGridViewTextBoxColumn { HeaderText = "File", ReadOnly = true, FillWeight = 22 };
        _colOpen = new DataGridViewButtonColumn { HeaderText = "Open", FillWeight = 7 };
        _colNotebook = new DataGridViewComboBoxColumn { HeaderText = "Notebook", FillWeight = 16 };
        _colTarget = new DataGridViewComboBoxColumn { HeaderText = "Target (section or section-group)", FillWeight = 30 };
        _colTitle = new DataGridViewTextBoxColumn { HeaderText = "Title (page becomes 'yyyy-MM <title>')", FillWeight = 16 };
        _colKind = new DataGridViewTextBoxColumn { HeaderText = "Kind", ReadOnly = true, FillWeight = 9 };
        _colStatus = new DataGridViewTextBoxColumn { HeaderText = "Status", ReadOnly = true, FillWeight = 11 };
        _grid.Columns.AddRange(_colEnabled, _colInst, _colFile, _colOpen, _colNotebook, _colTarget, _colTitle, _colKind, _colStatus);
        _grid.CellClick += Grid_CellClick;
        _grid.CellBeginEdit += Grid_CellBeginEdit;
        _grid.CellValueChanged += Grid_CellValueChanged;
        _grid.CellEndEdit += Grid_CellEndEdit;
        _grid.CellValidating += Grid_CellValidating;
        _grid.DataError += (_, e) => { e.ThrowException = false; LogGrid("DataError: " + (e.Exception?.Message ?? "?")); };
        // DIAGNOSTIC only (no commit) — finding out why a pick from the dropdown
        // does not stick. Logs the event; does not interfere with the grid.
        _grid.CurrentCellDirtyStateChanged += (_, __) =>
        {
            string colName = "?";
            var c = _grid.CurrentCell;
            if (c != null && c.OwningColumn != null) colName = c.OwningColumn.HeaderText ?? "?";
            LogGrid($"CurrentCellDirtyStateChanged: IsCurrentCellDirty={_grid.IsCurrentCellDirty} col={colName}");
        };

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
        RefreshTargetOptions();    // all targets; the per-row Notebook column filters each dropdown
        RefreshNotebookOptions(); // all notebooks, for the per-row Notebook column

        var all = BuildAllOptions();

        _grid.Rows.Clear();
        var files = _files
            .Where(f => !f.Error && f.Statements.Count > 0)
            .OrderBy(f => string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution, StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.FileName, StringComparer.OrdinalIgnoreCase);

        foreach (var f in files)
        {
            string inst = string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution;
            // Per-institution mapping: target + notebook.
            TargetOption? opt = null;
            string nb = "";
            if (_map.Institutions.TryGetValue(inst, out var e) && !string.IsNullOrWhiteSpace(e.Name))
            {
                // Match by Name + Kind + Parent so a same-name section in a different
                // group resolves to the mapped one.
                opt = all.FirstOrDefault(o => o.Name.Equals(e.Name, StringComparison.OrdinalIgnoreCase)
                    && o.Kind == e.Kind
                    && o.Parent.Equals(e.Parent ?? "", StringComparison.OrdinalIgnoreCase));
                if (opt == null)
                    opt = new TargetOption(e.Name, e.Kind == "group" ? "group" : "section", e.Notebook, e.Parent);
                nb = e.Notebook;
            }
            // Per-file settings: enabled + title (default title = institution).
            bool enabled = true;
            string title = inst;
            if (_map.Files.TryGetValue(f.FileName, out var fs))
            {
                enabled = fs.Enabled;
                title = string.IsNullOrWhiteSpace(fs.Title) ? inst : fs.Title;
            }
            int r = _grid.Rows.Add(enabled, inst, f.FileName, "Open", nb, opt, title,
                                  opt?.Kind ?? "",
                                  opt == null ? "UNMAPPED" : "mapped");
            _grid.Rows[r].DefaultCellStyle.BackColor = opt == null ? Color.LightYellow : Color.White;
        }
        _updating = false;
    }

    /// <summary>
    /// Sets the Target column's dropdown list to the targets of <c>notebook</c>
    /// (every notebook when <c>notebook</c> is empty). Includes mapped targets even
    /// if they don't exist in OneNote yet (created on run). Keeps combo values
    /// valid.
    /// </summary>
    private void SetTargetItems(string? notebook)
    {
        var list = new List<TargetOption>();
        if (_hierarchy != null)
            list.AddRange(BuildAllOptions());
        foreach (var kv in _map.Institutions)
        {
            var e = kv.Value;
            if (string.IsNullOrWhiteSpace(e.Name)) continue;
            var kind = e.Kind == "group" ? "group" : "section";
            if (!list.Any(o => string.Equals(o.Name, e.Name, StringComparison.OrdinalIgnoreCase) && o.Kind == kind
                && string.Equals(o.Parent, e.Parent ?? "", StringComparison.OrdinalIgnoreCase)))
                list.Add(new TargetOption(e.Name, kind, e.Notebook, e.Parent));
        }
        if (!string.IsNullOrEmpty(notebook))
            list = list.Where(o => o.Notebook == notebook).ToList();
        SortOptions(list);
        _colTarget.Items.Clear();
        _colTarget.Items.AddRange(list.ToArray());
    }

    /// <summary>Restore the Target dropdown to the full list (all notebooks).</summary>
    private void RefreshTargetOptions() => SetTargetItems(null);

    private static void SortOptions(List<TargetOption> list)
    {
        // Groups first, then by name — so the kind is easy to scan.
        list.Sort((a, b) =>
        {
            int c = string.CompareOrdinal(a.Kind, b.Kind);
            return c != 0 ? c : string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        });
    }

    /// <summary>Fill the per-row Notebook dropdown with all known notebooks.</summary>
    private void RefreshNotebookOptions()
    {
        _colNotebook.Items.Clear();
        var nbs = AllNotebooks();
        if (nbs.Count > 0) _colNotebook.Items.AddRange(nbs.ToArray());
    }

    /// <summary>All distinct notebook names in the hierarchy.</summary>
    private List<string> AllNotebooks()
    {
        var nbs = new List<string>();
        if (_hierarchy != null)
        {
            nbs.AddRange(_hierarchy.Groups.Select(g => g.Notebook));
            nbs.AddRange(_hierarchy.Sections.Select(s => s.Notebook));
        }
        return nbs.Where(s => !string.IsNullOrWhiteSpace(s))
                 .Distinct(StringComparer.OrdinalIgnoreCase)
                 .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                 .ToList();
    }

    /// <summary>The Notebook value of a grid row ("" if the row is invalid).</summary>
    private string GetRowNotebook(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _grid.Rows.Count) return "";
        return _grid.Rows[rowIndex].Cells[_colNotebook.Index].Value?.ToString() ?? "";
    }

    /// <summary>All section / section-group options across every notebook (recycle bins excluded).</summary>
    private List<TargetOption> BuildAllOptions()
    {
        var list = new List<TargetOption>();
        if (_hierarchy == null) return list;
        foreach (var g in _hierarchy.Groups)
            if (!g.Name.StartsWith("OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase))
                list.Add(new TargetOption(g.Name, "group", g.Notebook, g.Parent));
        foreach (var s in _hierarchy.Sections)
            if (!s.Name.StartsWith("OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase))
                list.Add(new TargetOption(s.Name, "section", s.Notebook, s.Parent));
        return list;
    }

    /// <summary>
    /// Normalize a Target cell value to a TargetOption. The combo cell can sometimes
    /// hold the display string ("[SECTION] Name") instead of the TargetOption object
    /// (a "value is not valid" state). A string is resolved back to a matching option
    /// so the value is valid again and the mapping is preserved.
    /// </summary>
    private TargetOption? ResolveTarget(object? value)
    {
        if (value is TargetOption o) return o;
        if (value is string s && s.Length > 0)
        {
            var all = BuildAllOptions();
            // Match the display text ("[SECTION] Name" / "[GROUP]   Name").
            foreach (var opt in all)
                if (string.Equals(opt.ToString(), s, StringComparison.Ordinal))
                    return opt;
            // Fallback: the string is a bare name.
            foreach (var opt in all)
                if (string.Equals(opt.Name, s, StringComparison.OrdinalIgnoreCase))
                    return opt;
        }
        return null;
    }

    /// <summary>
    /// Per-row filter: when the user opens the Target dropdown, show only the
    /// targets belonging to that row's Notebook.
    /// </summary>
    private void Grid_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
    {
        if (e.RowIndex >= 0 && e.ColumnIndex == _colTarget.Index)
            SetTargetItems(GetRowNotebook(e.RowIndex));
    }

    private void Grid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
    {
        // After editing, restore the full Target list so every row's value is valid again.
        if (e.ColumnIndex == _colTarget.Index)
            SetTargetItems(null);
        LogGrid($"CellEndEdit row={e.RowIndex} col={e.ColumnIndex} value={(e.RowIndex >= 0 ? (_grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "(null)") : "?")}");
    }

    /// <summary>
    /// The "Open" button next to the file name: shell-executes (opens) the PDF so
    /// the user can inspect it in detail.
    /// </summary>
    private void Grid_CellClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex != _colOpen.Index) return;
        string file = _grid.Rows[e.RowIndex].Cells[_colFile.Index].Value?.ToString() ?? "";
        if (file.Length == 0) return;
        var f = _files.FirstOrDefault(x => string.Equals(x.FileName, file, StringComparison.OrdinalIgnoreCase));
        if (f == null) return;
        string pdf = Path.Combine(_outDir, f.Category, f.FileName);
        if (!File.Exists(pdf))
        {
            Log("  Open: file not found at " + pdf, error: true);
            return;
        }
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdf) { UseShellExecute = true });
            Log("  Open: " + pdf);
        }
        catch (Exception ex)
        {
            Log("  Open failed: " + ex.Message, error: true);
        }
    }

    private void Grid_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
    {
        LogGrid($"CellValidating row={e.RowIndex} col={e.ColumnIndex} value={(e.RowIndex >= 0 ? (_grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "(null)") : "?")}");
    }

    private void Grid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || _updating) return;
        var row = _grid.Rows[e.RowIndex];

        if (e.ColumnIndex == _colNotebook.Index)
        {
            // Notebook changed: clear a target that no longer belongs to this notebook
            // (the per-row filter would make it an invalid pick).
            string nb = row.Cells[_colNotebook.Index].Value?.ToString() ?? "";
            var opt = ResolveTarget(row.Cells[_colTarget.Index].Value);
            _updating = true;
            try
            {
                if (opt != null && !string.IsNullOrEmpty(nb) && opt.Notebook != nb)
                {
                    row.Cells[_colTarget.Index].Value = null;
                    row.Cells[_colKind.Index].Value = "";
                    row.Cells[_colStatus.Index].Value = "UNMAPPED";
                    row.DefaultCellStyle.BackColor = Color.LightYellow;
                    LogGrid($"row={e.RowIndex}: notebook -> '{nb}'; target '{opt.Name}' is in '{opt.Notebook}' so it was cleared");
                }
            }
            finally { _updating = false; }
        }
        else if (e.ColumnIndex == _colTarget.Index)
        {
            var raw = row.Cells[_colTarget.Index].Value;
            var opt = ResolveTarget(raw);
            _updating = true;
            try
            {
                // Self-heal: if the grid stored a display string, write back a real
                // TargetOption so the value is valid again.
                if (raw is string && opt != null)
                    row.Cells[_colTarget.Index].Value = opt;
                row.Cells[_colKind.Index].Value = opt?.Kind ?? "";
                row.Cells[_colStatus.Index].Value = opt != null ? "mapped" : "UNMAPPED";
                row.DefaultCellStyle.BackColor = opt != null ? Color.White : Color.LightYellow;
                // Auto-sync the Notebook column to the picked target's notebook.
                if (opt != null && !string.IsNullOrEmpty(opt.Notebook))
                    row.Cells[_colNotebook.Index].Value = opt.Notebook;
            }
            catch (Exception ex)
            {
                LogGrid($"Grid_CellValueChanged error: {ex.Message}");
            }
            finally { _updating = false; }
        }
    }

    /// <summary>Run an action on the UI thread (fire-and-forget, guarded).</summary>
    private void OnUi(Action a)
    {
        try
        {
            if (_log == null || !_log.IsHandleCreated) return;
            if (_log.InvokeRequired) _log.BeginInvoke(a);
            else a();
        }
        catch { /* form already closing */ }
    }

    /// <summary>Per-file run settings read from the grid.</summary>
    private sealed class RunRow
    {
        public string FileName = "";
        public string Institution = "";
        public bool Enabled = true;
        public TargetOption? Target;
        public string Notebook = "";
        public string Title = "";
    }

    /// <summary>
    /// Read per-file run settings (enabled, target, notebook, title) from the grid.
    /// The grid is a UI control, so it must be read on the UI thread; this
    /// marshals a blocking read and returns plain data the background thread can use.
    /// </summary>
    private List<RunRow> SnapshotRunRows()
    {
        var rows = new List<RunRow>();
        try
        {
            if (_log == null || !_log.IsHandleCreated) return rows;
            Action fill = () =>
            {
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    rows.Add(new RunRow
                    {
                        FileName = row.Cells[_colFile.Index].Value?.ToString() ?? "",
                        Institution = row.Cells[_colInst.Index].Value?.ToString() ?? "",
                        Enabled = (bool)(row.Cells[_colEnabled.Index].Value ?? true),
                        Target = ResolveTarget(row.Cells[_colTarget.Index].Value),
                        Notebook = row.Cells[_colNotebook.Index].Value?.ToString() ?? "",
                        Title = row.Cells[_colTitle.Index].Value?.ToString() ?? "",
                    });
                }
            };
            if (_log.InvokeRequired) _log.Invoke(fill);
            else fill();
        }
        catch { /* form already closing — return what we have */ }
        return rows;
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

    private void LogGrid(string msg)
    {
        try
        {
            var s = "[grid] " + msg;
            if (_log != null && _log.IsHandleCreated && !_log.IsDisposed)
            {
                if (_log.InvokeRequired)
                    _log.BeginInvoke(new Action(() => AppendLogLine(s)));
                else
                    AppendLogLine(s);
            }
            else
            {
                _pendingLog.Add(s + "|");
            }
        }
        catch { }
    }

    private void AppendLogLine(string line)
    {
        _log.AppendText(line + Environment.NewLine);
        _log.ScrollToCaret();
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
        OnUi(() => { RefreshGrid(); RefreshInfo(); });
    }

    private void DoRun(bool dryRun)
    {
        if (_files.Count == 0)
        {
            Log("No statement files loaded — run StatementOrganizer first.", error: true);
            return;
        }

        // The grid is a UI control; read it on the UI thread into plain data.
        var rows = SnapshotRunRows();

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

        _cts = new System.Threading.CancellationTokenSource();
        if (!dryRun) OnUi(() => _btnStop.Enabled = true);

        int made = 0, skipped = 0, noTarget = 0, disabled = 0;
        int dpi = Core.GetDpi(_env);
        int maxPages = Core.GetMaxPages(_env);
        Log($"Importing {rows.Count} file(s) {(dryRun ? "(dry run)" : "into OneNote")} — one page per file");

        foreach (var rr in rows)
        {
            if (_cts is { IsCancellationRequested: true })
            {
                Log("Stopped by user — aborting remaining files.", error: true);
                break;
            }
            var f = _files.FirstOrDefault(x => string.Equals(x.FileName, rr.FileName, StringComparison.OrdinalIgnoreCase));
            if (f == null || f.Error || f.Statements.Count == 0) continue;

            if (!rr.Enabled)
            {
                Log($"  {f.FileName}: skipped (disabled).");
                disabled++;
                continue;
            }
            if (rr.Target == null)
            {
                Log($"  {f.FileName}: NO TARGET SET — skipping.", error: true);
                noTarget++;
                skipped++;
                continue;
            }

            string pdf = Path.Combine(_outDir, f.Category, f.FileName);
            if (!File.Exists(pdf))
            {
                Log($"  ! {f.FileName}: PDF not found at {pdf} — skipping.", error: true);
                skipped++;
                continue;
            }

            int year = f.Statements.Select(s => s.StatementDate?.Year ?? 0).Max();
            if (year == 0) year = DateTime.Now.Year;
            string pageName = Core.PageTitle(f, rr.Title);

            if (dryRun)
            {
                string label = rr.Target.Kind == "group" ? $"{rr.Target.Name} {year}" : rr.Target.Name;
                Log($"  [plan] {f.FileName} ({f.Statements.Count} account(s)) -> {label} / {pageName}");
                continue;
            }

            Log($"  {f.FileName} ({f.Statements.Count} account(s))");
            try
            {
                var entry = new MapEntry { Kind = rr.Target.Kind, Name = rr.Target.Name };
                var sec = ResolveSection(onenote!, ref h!, entry, year, rr.Notebook, rr.Title, rr.Target.Parent);
                Log($"    -> [{sec.Notebook}] {sec.Name} / {pageName}");
                string pageId = onenote!.CreateNote(sec.ObjectID, pageName);
                // Title line (becomes the OneNote page title), then a table with one row per account.
                var accRows = Core.BuildAccountRows(f);
                var rasters = PdfTools.Rasterize(pdf, dpi, maxPages);
                onenote.CommitPageFull(pageId, pageName, accRows, pdf, rasters, dpi);
                made++;
                Log($"    created + content set ({rasters.Count} image(s), {f.Statements.Count} account(s)).");
            }
            catch (Exception e)
            {
                Log("    FAILED: " + e.Message, error: true);
                if (e is System.Runtime.InteropServices.COMException ce)
                    Log("      HRESULT 0x" + ce.HResult.ToString("X8") + " (0x80042004 = OneNote write lock — try the local 'StatementsAll' notebook or reset OneDrive sync)", error: true);
                skipped++;
            }
        }

        OnUi(() => _btnStop.Enabled = false);
        _cts?.Dispose(); _cts = null;
        onenote?.Dispose();

        if (dryRun)
        {
            Log($"Dry run complete. {noTarget} file(s) had no target set, {disabled} disabled.");
        }
        else
        {
            Log($"Done. {made} OneNote page(s) created, {skipped} skipped, {disabled} disabled.");
            OnUi(() => SaveMappingFromGrid(silent: true));
            Log("OneNote is open so you can browse the new pages.");
        }
        OnUi(RefreshInfo);
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
        string pageId = onenote.CreateNote(sec.ObjectID, pageName);

        var rows = new List<PageXml.AccountRow>
        {
            new PageXml.AccountRow
            {
                Date = "2026-06-30",
                Name = "Example Checking (\u20261234)",
                Balance = "$26,372.47",
                Notes = new List<string>
                {
                    "Monthly statement - issued June 30, 2026",
                    "Opening balance: $23,056.08",
                }
            }
        };

        // Generated gradient image so the printout path is exercised even without a PDF.
        int w = 600, ht = 200;
        var px = new byte[w * ht * 3];
        for (int y = 0; y < ht; y++)
            for (int x = 0; x < w; x++)
            {
                int i = (y * w + x) * 3;
                px[i] = (byte)(x * 255 / w);
                px[i + 1] = (byte)(y * 255 / ht);
                px[i + 2] = 180;
            }
        string imgB64 = Convert.ToBase64String(PngWriter.Encode(px, w, ht));

        string? pdf = Core.FindFirstPdf(_outDir);
        string? fileB64 = null;
        if (pdf != null)
        {
            Log("Using real PDF for the test: " + pdf);
            rows[0].Notes.Add("PDF: " + pdf);
            fileB64 = Convert.ToBase64String(File.ReadAllBytes(pdf));
            var pages = PdfTools.Rasterize(pdf, Core.GetDpi(_env), 1);
            if (pages.Count > 0) imgB64 = Convert.ToBase64String(pages[0].Png);
        }

        Log("Committing (core always; binary is best-effort) ...");
        var (imgOk, fileOk, note) = onenote.TryCommitPageWithBinary(
            pageId, pageName, rows, imgB64, w, ht, fileB64,
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
    /// Resolve (and create, best-effort) the target section for a mapping entry.
    /// For a group mapping, the section is "{Title} {year}" (created in the group if needed).
    /// Section targets are disambiguated by their containing group (Parent), because the
    /// same section name can exist in different groups.
    /// </summary>
    private static SectionInfo ResolveSection(OneNote onenote, ref Hierarchy h, MapEntry entry, int year, string notebook, string title, string parent)
    {
        if (entry.Kind == "group")
        {
            string notebookId = NotebookId(h, notebook);
            var g = h.Groups.FirstOrDefault(x => x.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase) && SameParent(x.Parent, notebookId))
                  ?? h.Groups.FirstOrDefault(x => x.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase))
                  ?? throw new Exception($"Section group \"{entry.Name}\" no longer exists in OneNote.");
            string baseName = string.IsNullOrWhiteSpace(title) ? entry.Name : title.Trim();
            string name = $"{baseName} {year}";

            // The section must live inside THIS group (Parent = group ObjectID).
            var existing = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, g.ObjectID));
            if (existing != null) return existing;

            // Best-effort create. Section creation is often blocked (0x80042004);
            // if it fails, give a clear, actionable message.
            try
            {
                onenote.CreateSectionInGroup(name, g);
                h = onenote.GetHierarchy();
            }
            catch (Exception e)
            {
                throw new Exception($"Could not create section \"{name}\" in group \"{entry.Name}\" ({e.Message}). Create it manually in OneNote and re-run.");
            }
            var created = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, g.ObjectID));
            if (created != null) return created;
            throw new Exception($"Section \"{name}\" was not created in group \"{entry.Name}\". Create it manually in OneNote and re-run.");
        }

        // Section target — disambiguate by containing group (Parent).
        var sec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, parent));
        if (sec != null) return sec;

        // Not present: create it. If the target is inside a group (parent set), create it
        // there; otherwise create it directly under the notebook.
        try
        {
            if (string.IsNullOrEmpty(parent))
                onenote.CreateSection(entry.Name, string.IsNullOrEmpty(notebook) ? secNotebookFallback(h) : notebook);
            else
            {
                var g = h.Groups.FirstOrDefault(x => SameParent(x.ObjectID, parent))
                      ?? throw new Exception($"Section group (ObjectID {parent}) not found (needed to create section \"{entry.Name}\").");
                onenote.CreateSectionInGroup(entry.Name, g);
            }
            h = onenote.GetHierarchy();
        }
        catch (Exception e)
        {
            throw new Exception($"Could not create section \"{entry.Name}\" ({e.Message}). Create it manually in OneNote and re-run.");
        }
        var createdSec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, parent));
        if (createdSec != null) return createdSec;
        throw new Exception($"Section \"{entry.Name}\" was not created. Create it manually in OneNote and re-run.");
    }

    static bool SameParent(string? a, string? b)
    {
        a = string.IsNullOrEmpty(a) ? "" : a!;
        b = string.IsNullOrEmpty(b) ? "" : b!;
        return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Resolve a notebook name to its ObjectID ("" if not in the hierarchy).</summary>
    static string NotebookId(Hierarchy h, string name)
    {
        if (string.IsNullOrEmpty(name)) return "";
        return h.Notebooks.FirstOrDefault(n => n.Name.Equals(name, StringComparison.OrdinalIgnoreCase))?.ObjectID ?? "";
    }

    static string secNotebookFallback(Hierarchy h) => h.Sections.FirstOrDefault()?.Notebook ?? "";

    /// <summary>
    /// Toggle the "On" column of every file row: if all are ON, turn all OFF;
    /// otherwise turn all ON.
    /// </summary>
    private void ToggleAll()
    {
        if (_grid.Rows.Count == 0) return;
        bool allOn = true;
        foreach (DataGridViewRow r in _grid.Rows)
            if (!(r.Cells[_colEnabled.Index].Value is bool b && b)) { allOn = false; break; }
        _updating = true;
        try
        {
            foreach (DataGridViewRow r in _grid.Rows)
                r.Cells[_colEnabled.Index].Value = !allOn;
            Log($"All file rows set to {(allOn ? "OFF" : "ON")} ({_grid.Rows.Count} rows).");
        }
        finally { _updating = false; }
    }

    private void SaveMappingFromGrid(bool silent = false)
    {
        var map = new OneMap();
        var seenInst = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (DataGridViewRow row in _grid.Rows)
        {
            string inst = row.Cells[_colInst.Index].Value?.ToString() ?? "";
            string file = row.Cells[_colFile.Index].Value?.ToString() ?? "";
            string nb = row.Cells[_colNotebook.Index].Value?.ToString() ?? "";
            // Per-file settings: enabled + title.
            if (file.Length > 0)
            {
                map.Files[file] = new FileSetting
                {
                    Enabled = (bool)(row.Cells[_colEnabled.Index].Value ?? true),
                    Title = row.Cells[_colTitle.Index].Value?.ToString() ?? "",
                };
            }
            // Per-institution mapping: target + notebook (first file for the institution wins).
            var opt = ResolveTarget(row.Cells[_colTarget.Index].Value);
            if (opt != null && !seenInst.Contains(inst))
            {
                seenInst.Add(inst);
                map.Institutions[inst] = new MapEntry
                {
                    Kind = opt.Kind,
                    Name = opt.Name,
                    Notebook = string.IsNullOrEmpty(nb) ? opt.Notebook : nb,
                    Parent = opt.Parent,
                };
            }
        }
        try
        {
            string mapPath = Core.MapPath(_root);
            Core.SaveMap(mapPath, map);
            if (!silent) Log($"Saved {map.Institutions.Count} institution mapping(s) and {map.Files.Count} file setting(s) to {mapPath}");
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
