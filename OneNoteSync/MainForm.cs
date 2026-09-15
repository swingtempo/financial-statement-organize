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
    private bool _rebuildingOptions;

    // Controls.
    private Panel _top = null!;
    private Button _btnReload = null!;
    private Button _btnDryRun = null!;
    private Button _btnTest = null!;
    private Button _btnRun = null!;
    private Button _btnSave = null!;
    private Button _btnClear = null!;
    private ComboBox _cboNotebook = null!;
    private Label _info = null!;
    private DataGridView _grid = null!;
    private DataGridViewComboBoxColumn _colTarget = null!;
    private DataGridViewTextBoxColumn _colInst = null!;
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
        public TargetOption(string name, string kind, string notebook)
        {
            Name = name; Kind = kind; Notebook = notebook;
        }
        public override string ToString() => Kind == "group" ? $"[GROUP]   {Name}" : $"[SECTION] {Name}";
        // Value equality (Name + Kind). The combo cell holds an instance built by one
        // list and the Items hold instances built by another; reference equality would
        // make the grid report "value is not valid". Comparing by Name + Kind fixes that.
        public override bool Equals(object? obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (obj is not TargetOption other) return false;
            return string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase) && Kind == other.Kind;
        }
        public override int GetHashCode()
        {
            unchecked
            {
                int h = (Name ?? "").GetHashCode();
                h = (h * 397) ^ (Kind ?? "").GetHashCode();
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

        // Top action bar (two rows: buttons, then the notebook filter).
        _top = new Panel { Dock = DockStyle.Top, Height = 84 };
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

        // Notebook filter (row 2) — filters the Target dropdown by notebook.
        var lblNb = new Label { Text = "Notebook:", Location = new Point(10, 51), AutoSize = true };
        _top.Controls.Add(lblNb);
        _cboNotebook = new ComboBox { Location = new Point(84, 43), Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };
        _cboNotebook.SelectedIndexChanged += Notebook_Changed;
        _top.Controls.Add(_cboNotebook);

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
        RefreshNotebookFilter();
        RefreshTargetOptions();

        // Master lookup (all notebooks) so a saved target that belongs to a
        // different notebook than the current filter still resolves.
        var all = BuildAllOptions();

        _grid.Rows.Clear();
        var institutions = _files
            .Where(f => !f.Error && f.Statements.Count > 0)
            .Select(f => string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase);

        foreach (var inst in institutions)
        {
            TargetOption? opt = null;
            if (_map.Institutions.TryGetValue(inst, out var e) && !string.IsNullOrWhiteSpace(e.Name))
            {
                opt = all.FirstOrDefault(o => o.Name.Equals(e.Name, StringComparison.OrdinalIgnoreCase) && o.Kind == e.Kind);
                if (opt == null)
                    opt = new TargetOption(e.Name, e.Kind == "group" ? "group" : "section", "");
            }
            int r = _grid.Rows.Add(inst, opt, opt?.Kind ?? "",
                                   opt == null ? "UNMAPPED" : "mapped");
            // Highlight unmapped rows.
            _grid.Rows[r].DefaultCellStyle.BackColor = opt == null ? Color.LightYellow : Color.White;
        }
        ClearOutOfFilterCells();
        _updating = false;
    }

    /// <summary>
    /// Rebuild the Target dropdown: sections / section-groups of the currently
    /// selected notebook only (true filter). Empty when the hierarchy is not loaded
    /// or no notebook is selected.
    /// </summary>
    private void RefreshTargetOptions()
    {
        _colTarget.Items.Clear();
        string nb = SelectedNotebook();
        var list = new List<TargetOption>();
        if (_hierarchy != null && nb != "")
            list.AddRange(BuildAllOptions().Where(o => o.Notebook == nb));
        // Ensure mapped targets are valid options even if they don't exist in
        // OneNote yet (they will be created on run). This keeps the combo cell
        // value valid (and avoids the "value is not valid" DataError).
        if (nb != "")
        {
            foreach (var kv in _map.Institutions)
            {
                var e = kv.Value;
                if (string.IsNullOrWhiteSpace(e.Name)) continue;
                var kind = e.Kind == "group" ? "group" : "section";
                if (!list.Any(o => string.Equals(o.Name, e.Name, StringComparison.OrdinalIgnoreCase) && o.Kind == kind))
                    list.Add(new TargetOption(e.Name, kind, nb));
            }
        }
        SortOptions(list);
        _colTarget.Items.AddRange(list.ToArray());
    }

    private static void SortOptions(List<TargetOption> list)
    {
        // Groups first, then by name — so the kind is easy to scan.
        list.Sort((a, b) =>
        {
            int c = string.CompareOrdinal(a.Kind, b.Kind);
            return c != 0 ? c : string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        });
    }

    private void Notebook_Changed(object? sender, EventArgs e)
    {
        if (_rebuildingOptions) return;
        ApplyNotebookFilter();
    }

    /// <summary>Rebuild the dropdown and clear any picks that fall outside the selected notebook.</summary>
    private void ApplyNotebookFilter()
    {
        RefreshTargetOptions();
        ClearOutOfFilterCells();
    }

    /// <summary>Build the notebook list (no "All"), preserving the prior selection where possible.</summary>
    private void RefreshNotebookFilter()
    {
        string prev = _cboNotebook.SelectedItem?.ToString() ?? "";
        _rebuildingOptions = true;
        _cboNotebook.Items.Clear();
        if (_hierarchy != null)
        {
            var nbs = new List<string>();
            nbs.AddRange(_hierarchy.Groups.Select(g => g.Notebook));
            nbs.AddRange(_hierarchy.Sections.Select(s => s.Notebook));
            foreach (var nb in nbs.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
                _cboNotebook.Items.Add(nb);
        }
        int idx = -1;
        for (int i = 0; i < _cboNotebook.Items.Count; i++)
            if (_cboNotebook.Items[i]!.ToString() == prev) { idx = i; break; }
        _cboNotebook.SelectedIndex = idx >= 0 ? idx : (_cboNotebook.Items.Count > 0 ? 0 : -1);
        _rebuildingOptions = false;
    }

    /// <summary>The selected notebook's name; "" when nothing is selected.</summary>
    private string SelectedNotebook()
    {
        return _cboNotebook.SelectedItem?.ToString() ?? "";
    }

    /// <summary>
    /// Clear any target picks that no longer fall inside the selected notebook
    /// (the dropdown is a true filter, so out-of-filter picks would be invalid).
    /// </summary>
    private void ClearOutOfFilterCells()
    {
        string nb = SelectedNotebook();
        bool wasUpdating = _updating;
        _updating = true;
        foreach (DataGridViewRow row in _grid.Rows)
        {
            var opt = row.Cells[_colTarget.Index].Value as TargetOption;
            if (opt != null && opt.Notebook != nb)
            {
                row.Cells[_colTarget.Index].Value = null;
                row.Cells[_colKind.Index].Value = "";
                row.Cells[_colStatus.Index].Value = "UNMAPPED";
                row.DefaultCellStyle.BackColor = Color.LightYellow;
            }
        }
        _updating = wasUpdating;
    }

    /// <summary>All section / section-group options across every notebook.</summary>
    private List<TargetOption> BuildAllOptions()
    {
        var list = new List<TargetOption>();
        if (_hierarchy == null) return list;
        foreach (var g in _hierarchy.Groups) list.Add(new TargetOption(g.Name, "group", g.Notebook));
        foreach (var s in _hierarchy.Sections) list.Add(new TargetOption(s.Name, "section", s.Notebook));
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

    private void Grid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
    {
        LogGrid($"CellEndEdit row={e.RowIndex} col={e.ColumnIndex} value={(e.RowIndex >= 0 ? (_grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "(null)") : "?")}");
    }

    private void Grid_CellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
    {
        LogGrid($"CellValidating row={e.RowIndex} col={e.ColumnIndex} value={(e.RowIndex >= 0 ? (_grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "(null)") : "?")}");
    }

    private void Grid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        if (e.ColumnIndex != _colTarget.Index) return;
        if (_updating) return; // prevent re-entrancy from our own writes

        var raw = _grid.Rows[e.RowIndex].Cells[_colTarget.Index].Value;
        if (raw is not TargetOption)
            LogGrid($"[diag] row={e.RowIndex} Target value is NOT a TargetOption: type={(raw?.GetType().Name ?? "null")} text={(raw?.ToString() ?? "(null)")}");
        var opt = ResolveTarget(raw);

        _updating = true;
        try
        {
            // Self-heal: if the grid stored a display string, write back a real
            // TargetOption so the value is valid again (fixes "value is not valid"
            // and the empty Kind/Status).
            if (raw is string && opt != null)
                _grid.Rows[e.RowIndex].Cells[_colTarget.Index].Value = opt;
            _grid.Rows[e.RowIndex].Cells[_colKind.Index].Value = opt?.Kind ?? "";
            bool mapped = opt != null;
            _grid.Rows[e.RowIndex].Cells[_colStatus.Index].Value = mapped ? "mapped" : "UNMAPPED";
            _grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = mapped ? Color.White : Color.LightYellow;
        }
        catch (Exception ex)
        {
            LogGrid($"Grid_CellValueChanged error: {ex.Message}");
        }
        finally
        {
            _updating = false;
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

    /// <summary>
    /// Read the grid's institution -&gt; target mapping. The grid is a UI control, so
    /// it must be read on the UI thread; this marshals a blocking read and returns
    /// a plain dictionary the background thread can safely use.
    /// </summary>
    private Dictionary<string, MapEntry> SnapshotMappings()
    {
        var maps = new Dictionary<string, MapEntry>(StringComparer.OrdinalIgnoreCase);
        try
        {
            if (_log == null || !_log.IsHandleCreated) return maps;
            Action fill = () =>
            {
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    string inst = row.Cells[_colInst.Index].Value?.ToString() ?? "";
                    var opt = ResolveTarget(row.Cells[_colTarget.Index].Value);
                    if (opt == null) continue;
                    maps[inst] = new MapEntry { Kind = opt.Kind, Name = opt.Name };
                }
            };
            if (_log.InvokeRequired) _log.Invoke(fill);
            else fill();
        }
        catch { /* form already closing — return what we have */ }
        return maps;
    }

    /// <summary>Read the selected notebook name on the UI thread.</summary>
    private string SnapshotNotebook()
    {
        string nb = "";
        try
        {
            if (_log == null || !_log.IsHandleCreated) return nb;
            if (_log.InvokeRequired) _log.Invoke(new Action(() => nb = SelectedNotebook()));
            else nb = SelectedNotebook();
        }
        catch { }
        return nb;
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

        // The grid is a UI control; read it on the UI thread into a plain dict.
        var maps = SnapshotMappings();

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
        string notebook = SnapshotNotebook();
        Log($"Importing {_files.Count} file(s) {(dryRun ? "(dry run)" : "into OneNote")}");

        foreach (var f in _files.Where(f => !f.Error && f.Statements.Count > 0))
        {
            string inst = string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution;
            if (!maps.TryGetValue(inst, out var entry))
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

            int dpi = Core.GetDpi(_env);
            int maxPages = Core.GetMaxPages(_env);

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
                try
                {
                    var sec = ResolveSection(onenote!, ref h!, entry, year, notebook);
                    Log($"    -> [{sec.Notebook}] {sec.Name} / {pageName}");
                    string pageId = onenote!.CreateNote(sec.Guid, pageName);
                    var lines = Core.BuildSummaryLines(f, st);
                    string summary = string.Join("\n", lines);
                    var rasters = PdfTools.Rasterize(pdf, dpi, maxPages);
                    onenote.CommitPageFull(pageId, summary, pdf, rasters, dpi);
                    made++;
                    Log($"    created + content set ({rasters.Count} image(s)).");
                }
                catch (Exception e)
                {
                    Log("    FAILED: " + e.Message, error: true);
                    if (e is System.Runtime.InteropServices.COMException ce)
                        Log("      HRESULT 0x" + ce.HResult.ToString("X8") + " (0x80042004 = OneNote write lock — try the local 'StatementsAll' notebook or reset OneDrive sync)", error: true);
                    skipped++;
                }
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
        string pageId = onenote.CreateNote(sec.Guid, pageName);

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
    private static SectionInfo ResolveSection(OneNote onenote, ref Hierarchy h, MapEntry entry, int year, string notebook)
    {
        if (entry.Kind == "group")
        {
            var g = h.Groups.FirstOrDefault(g => g.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase))
                  ?? throw new Exception($"Section group \"{entry.Name}\" no longer exists in OneNote.");
            string name = $"{entry.Name} {year}";

            var existing = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (existing != null) return existing;

            onenote.CreateSectionInGroup(name, g.Guid, g.Notebook);
            h = onenote.GetHierarchy();
            return h.Sections.First(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        var sec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
        if (sec != null) return sec;

        onenote.CreateSection(entry.Name, string.IsNullOrEmpty(notebook) ? secNotebookFallback(h) : notebook);
        h = onenote.GetHierarchy();
        return h.Sections.First(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase));
    }

    static string secNotebookFallback(Hierarchy h) => h.Sections.FirstOrDefault()?.Notebook ?? "";

    private void SaveMappingFromGrid(bool silent = false)
    {
        var map = new OneMap();
        foreach (DataGridViewRow row in _grid.Rows)
        {
            string inst = row.Cells[_colInst.Index].Value?.ToString() ?? "";
            var opt = ResolveTarget(row.Cells[_colTarget.Index].Value);
            if (opt == null) continue; // skip unmapped
            map.Institutions[inst] = new MapEntry { Kind = opt.Kind, Name = opt.Name };
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
