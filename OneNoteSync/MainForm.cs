using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OneNoteSync;

/// <summary>
/// Windows Forms front-end for OneNoteSync. Two grids:
///   - TOP:    institution -&gt; OneNote location mapping (Notebook / Target).
///   - BOTTOM: per-file settings (On / Title).
/// Plus a live debug log of everything the sync does.
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

    // TOP grid: institution -> notebook/target.
    private DataGridView _mapGrid = null!;
    private DataGridViewTextBoxColumn _mapColInst = null!;
    private DataGridViewComboBoxColumn _mapColNotebook = null!;
    private DataGridViewComboBoxColumn _mapColTarget = null!;

    // BOTTOM grid: per-file settings.
    private DataGridView _fileGrid = null!;
    private DataGridViewCheckBoxColumn _colEnabled = null!;
    private DataGridViewTextBoxColumn _colInst = null!;
    private DataGridViewTextBoxColumn _colFile = null!;
    private DataGridViewButtonColumn _colOpen = null!;
    private DataGridViewCheckBoxColumn _colDate = null!;
    private DataGridViewTextBoxColumn _colTitle = null!;
    private DataGridViewButtonColumn _colReset = null!;
    private DataGridViewTextBoxColumn _colStatus = null!;
    private int _editingTitleRow = -1;

    private RichTextBox _log = null!;
    private readonly List<string> _pendingLog = new();

    /// <summary>
    /// A pickable target. Carries the real section/section-group name, its kind,
    /// its owning notebook, and the containing group (both ObjectID, for internal
    /// disambiguation, and NAME, which is what gets persisted to onemap.json).
    /// </summary>
    private sealed class TargetOption
    {
        public string Name { get; }
        public string Kind { get; }        // "section" or "group"
        public string Notebook { get; }
        public string Parent { get; }      // containing group ObjectID (internal; "" if top-level)
        public string ParentName { get; }  // containing group NAME (persisted; "" if top-level)
        public TargetOption(string name, string kind, string notebook, string parent = "", string parentName = "")
        {
            Name = name; Kind = kind; Notebook = notebook; Parent = parent; ParentName = parentName;
        }
        public override string ToString()
        {
            if (Kind == "group")
            {
                string nb = string.IsNullOrEmpty(Notebook) ? "" : $"  ({Notebook})";
                return $"[GROUP] {Name}{nb}";
            }
            // Section: show the containing section-group name when present, else the notebook.
            if (!string.IsNullOrEmpty(ParentName))
                return $"[SECTION] {Name}  (in {ParentName})";
            string nbs = string.IsNullOrEmpty(Notebook) ? "" : $"  ({Notebook})";
            return $"[SECTION] {Name}{nbs}";
        }
        // Value equality (Name + Kind + Notebook + Parent). The combo cell holds an
        // instance built by one list and the Items hold instances built by another;
        // reference equality would make the grid report "value is not valid".
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
        // Route any Console.* output into the GUI log so nothing is lost (WinExe has no console).
        var cw = new ConsoleToLog(s => Log(s));
        Console.SetOut(cw);
        Console.SetError(cw);
        LoadData();
        // Auto-load OneNote at boot (per request).
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
        Height = 780;
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

        // TOP grid: institution -> notebook / target mapping.
        _mapGrid = new DataGridView
        {
            Dock = DockStyle.Top,
            Height = 170,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = SystemColors.Window,
        };
        _mapColInst = new DataGridViewTextBoxColumn { HeaderText = "Institution", ReadOnly = true, FillWeight = 22 };
        _mapColNotebook = new DataGridViewComboBoxColumn { HeaderText = "Notebook", FillWeight = 24 };
        _mapColTarget = new DataGridViewComboBoxColumn { HeaderText = "Target (section or section-group)", FillWeight = 54 };
        _mapGrid.Columns.AddRange(_mapColInst, _mapColNotebook, _mapColTarget);
        _mapGrid.CellBeginEdit += MapGrid_CellBeginEdit;
        _mapGrid.CellEndEdit += MapGrid_CellEndEdit;
        _mapGrid.CellValueChanged += MapGrid_CellValueChanged;
        _mapGrid.DataError += (_, e) => { e.ThrowException = false; LogGrid("mapGrid DataError: " + (e.Exception?.Message ?? "?")); };

        // BOTTOM grid: per-file settings.
        _fileGrid = new DataGridView
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
        _colFile = new DataGridViewTextBoxColumn { HeaderText = "File", ReadOnly = true, FillWeight = 24 };
        _colOpen = new DataGridViewButtonColumn { HeaderText = "Open", FillWeight = 7 };
        _colDate = new DataGridViewCheckBoxColumn { HeaderText = "Date", FillWeight = 6 };
        _colTitle = new DataGridViewTextBoxColumn { HeaderText = "Title", FillWeight = 26 };
        _colReset = new DataGridViewButtonColumn { HeaderText = "Reset", FillWeight = 8 };
        _colStatus = new DataGridViewTextBoxColumn { HeaderText = "Status", ReadOnly = true, FillWeight = 12 };
        _fileGrid.Columns.AddRange(_colEnabled, _colInst, _colFile, _colOpen, _colDate, _colTitle, _colReset, _colStatus);
        _fileGrid.CellClick += FileGrid_CellClick;
        _fileGrid.CellPainting += FileGrid_CellPaint;
        _fileGrid.CellBeginEdit += FileGrid_CellBeginEdit;
        _fileGrid.CellEndEdit += FileGrid_CellEndEdit;
        _fileGrid.CellValueChanged += FileGrid_CellValueChanged;
        _fileGrid.DataError += (_, e) => { e.ThrowException = false; LogGrid("fileGrid DataError: " + (e.Exception?.Message ?? "?")); };

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

        // Order matters for Dock layout. For Dock=Top, the control added LAST is
        // outermost (topmost), so add mapGrid, then info, then top. Fill goes last.
        Controls.Add(_fileGrid);
        Controls.Add(_log);
        Controls.Add(_mapGrid);
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
        RefreshMapGrid();
        RefreshFileGrid();
    }

    private static string InstOf(StatementFile f)
        => string.IsNullOrWhiteSpace(f.Institution) ? f.Category : f.Institution;

    /// <summary>Refresh the TOP grid: one row per institution (Notebook + Target).</summary>
    private void RefreshMapGrid()
    {
        _updating = true;
        RefreshNotebookOptions(_mapColNotebook);
        SetTargetItems(_mapColTarget, null);
        var all = BuildAllOptions();

        // Distinct institutions: from the files, plus any saved in the map.
        var insts = new List<string>();
        foreach (var f in _files)
        {
            string inst = InstOf(f);
            if (!insts.Contains(inst, StringComparer.OrdinalIgnoreCase)) insts.Add(inst);
        }
        foreach (var k in _map.Institutions.Keys)
            if (!insts.Contains(k, StringComparer.OrdinalIgnoreCase)) insts.Add(k);
        insts.Sort(StringComparer.OrdinalIgnoreCase);

        _mapGrid.Rows.Clear();
        foreach (string inst in insts)
        {
            TargetOption? opt = null;
            string nb = "";
            if (_map.Institutions.TryGetValue(inst, out var e) && !string.IsNullOrWhiteSpace(e.Name))
            {
                // Match by Name + Kind + (ParentName OR Notebook) so a saved target re-resolves
                // to the hierarchy option (carrying its real ObjectID) even if the persisted
                // ParentName string differs slightly.
                opt = all.FirstOrDefault(o =>
                    string.Equals(o.Name, e.Name, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(o.Kind, e.Kind, StringComparison.OrdinalIgnoreCase)
                    && (string.Equals(o.ParentName, e.Parent ?? "", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(o.Notebook, e.Notebook, StringComparison.OrdinalIgnoreCase)));
                if (opt == null)
                    opt = new TargetOption(e.Name, e.Kind == "group" ? "group" : "section", e.Notebook, "", e.Parent);
                nb = e.Notebook;
            }
            int r = _mapGrid.Rows.Add(inst, nb, opt);
            _mapGrid.Rows[r].DefaultCellStyle.BackColor = opt == null ? Color.LightYellow : Color.White;
        }
        _updating = false;
    }

    /// <summary>Refresh the BOTTOM grid: one row per file (On + Title + Status).</summary>
    private void RefreshFileGrid()
    {
        _updating = true;
        _fileGrid.Rows.Clear();
        var files = _files
            .Where(f => !f.Error && f.Statements.Count > 0)
            .OrderBy(f => InstOf(f), StringComparer.OrdinalIgnoreCase)
            .ThenBy(f => f.FileName, StringComparer.OrdinalIgnoreCase);

        foreach (var f in files)
        {
            string inst = InstOf(f);
            bool enabled = true;
            bool date = false;
            string title = Core.DefaultTitle(f);
            if (_map.Files.TryGetValue(f.FileName, out var fs))
            {
                enabled = fs.Enabled;
                date = fs.Date;
                title = string.IsNullOrWhiteSpace(fs.Title) ? Core.DefaultTitle(f) : fs.Title;
            }
            int r = _fileGrid.Rows.Add(enabled, inst, f.FileName, "Open", date, title, "Reset", "");
        }
        _updating = false;
    }

    /// <summary>
    /// Set a Target column's dropdown list to the targets of <c>notebook</c>
    /// (every notebook when <c>notebook</c> is empty). Includes mapped targets even
    /// if they don't exist in OneNote yet (created on run).
    /// </summary>
    private void SetTargetItems(DataGridViewComboBoxColumn col, string? notebook)
    {
        var list = new List<TargetOption>();
        if (_hierarchy != null)
            list.AddRange(BuildAllOptions());
        foreach (var kv in _map.Institutions)
        {
            var e = kv.Value;
            if (string.IsNullOrWhiteSpace(e.Name)) continue;
            var kind = e.Kind == "group" ? "group" : "section";
            // Add only if not already present. Match by Name+Kind+Notebook (display identity)
            // OR Name+Kind+ParentName — so a saved target that already exists (even with a
            // slightly different ParentName string) is not re-added once per mapped institution.
            bool present = list.Any(o =>
                string.Equals(o.Kind, kind, StringComparison.OrdinalIgnoreCase)
                && string.Equals(o.Name, e.Name, StringComparison.OrdinalIgnoreCase)
                && (string.Equals(o.Notebook, e.Notebook, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(o.ParentName, e.Parent ?? "", StringComparison.OrdinalIgnoreCase)));
            if (!present)
                list.Add(new TargetOption(e.Name, kind, e.Notebook, "", e.Parent));
        }
        if (!string.IsNullOrEmpty(notebook))
            list = list.Where(o => o.Notebook == notebook).ToList();
        // Safety net: never present two visually-identical entries (same Kind+Name+Notebook).
        // OneNote keeps section/group names unique within a notebook, so this key is unique.
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var deduped = new List<TargetOption>();
        foreach (var o in list)
            if (seen.Add(o.Kind + "|" + o.Name + "|" + (o.Notebook ?? "")))
                deduped.Add(o);
        SortOptions(deduped);
        col.Items.Clear();
        col.Items.AddRange(deduped.ToArray());
    }

    private static void SortOptions(List<TargetOption> list)
    {
        list.Sort((a, b) =>
        {
            int c = string.CompareOrdinal(a.Kind, b.Kind);
            return c != 0 ? c : string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
        });
    }

    /// <summary>Fill a per-row Notebook dropdown with all known notebooks.</summary>
    private void RefreshNotebookOptions(DataGridViewComboBoxColumn col)
    {
        col.Items.Clear();
        var nbs = AllNotebooks();
        if (nbs.Count > 0) col.Items.AddRange(nbs.ToArray());
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

    /// <summary>The Notebook value of a map-grid row ("" if the row is invalid).</summary>
    private string GetMapRowNotebook(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= _mapGrid.Rows.Count) return "";
        return _mapGrid.Rows[rowIndex].Cells[_mapColNotebook.Index].Value?.ToString() ?? "";
    }

    /// <summary>All section / section-group options across every notebook (recycle bins excluded).</summary>
    private List<TargetOption> BuildAllOptions()
    {
        var list = new List<TargetOption>();
        if (_hierarchy == null) return list;
        foreach (var g in _hierarchy.Groups)
            if (!g.Name.StartsWith("OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase))
                list.Add(new TargetOption(g.Name, "group", g.Notebook, g.Parent, g.Notebook));
        foreach (var s in _hierarchy.Sections)
            if (!s.Name.StartsWith("OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase))
                list.Add(new TargetOption(s.Name, "section", s.Notebook, s.Parent, s.ParentName));
        return list;
    }

    /// <summary>
    /// Normalize a Target cell value to a TargetOption. The combo cell can sometimes
    /// hold the display string ("[SECTION] Name") instead of the object (a "value is
    /// not valid" state). A string is resolved back to a matching option.
    /// </summary>
    private TargetOption? ResolveTarget(object? value)
    {
        if (value is TargetOption o) return o;
        if (value is string s && s.Length > 0)
        {
            var all = BuildAllOptions();
            foreach (var opt in all)
                if (string.Equals(opt.ToString(), s, StringComparison.Ordinal))
                    return opt;
            foreach (var opt in all)
                if (string.Equals(opt.Name, s, StringComparison.OrdinalIgnoreCase))
                    return opt;
        }
        return null;
    }

    // ------------------------------------------------------------------
    // Grid event handlers
    // ------------------------------------------------------------------

    /// <summary>Per-row filter: opening the Target dropdown shows only that row's Notebook targets.</summary>
    private void MapGrid_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
    {
        if (e.RowIndex >= 0 && e.ColumnIndex == _mapColTarget.Index)
            SetTargetItems(_mapColTarget, GetMapRowNotebook(e.RowIndex));
    }

    private void MapGrid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.ColumnIndex == _mapColTarget.Index)
            SetTargetItems(_mapColTarget, null);
    }

    private void MapGrid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || _updating) return;
        var row = _mapGrid.Rows[e.RowIndex];

        if (e.ColumnIndex == _mapColNotebook.Index)
        {
            // Notebook changed: clear a target that no longer belongs to this notebook.
            string nb = row.Cells[_mapColNotebook.Index].Value?.ToString() ?? "";
            var opt = ResolveTarget(row.Cells[_mapColTarget.Index].Value);
            _updating = true;
            try
            {
                if (opt != null && !string.IsNullOrEmpty(nb) && opt.Notebook != nb)
                    row.Cells[_mapColTarget.Index].Value = null;
            }
            finally { _updating = false; }
        }
        else if (e.ColumnIndex == _mapColTarget.Index)
        {
            var raw = row.Cells[_mapColTarget.Index].Value;
            var opt = ResolveTarget(raw);
            _updating = true;
            try
            {
                if (raw is string && opt != null)
                    row.Cells[_mapColTarget.Index].Value = opt;
                // Auto-sync the Notebook column to the picked target's notebook.
                if (opt != null && !string.IsNullOrEmpty(opt.Notebook))
                    row.Cells[_mapColNotebook.Index].Value = opt.Notebook;
            }
            catch (Exception ex)
            {
                LogGrid($"MapGrid_CellValueChanged error: {ex.Message}");
            }
            finally { _updating = false; }
        }
    }

    /// <summary>The "Open" button in the file grid: shell-executes (opens) the PDF.</summary>
    private void FileGrid_CellClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        if (e.ColumnIndex == _colReset.Index)
        {
            ResetFileRow(e.RowIndex);
            return;
        }
        if (e.ColumnIndex != _colOpen.Index) return;
        string file = _fileGrid.Rows[e.RowIndex].Cells[_colFile.Index].Value?.ToString() ?? "";
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

    /// <summary>Reset a file row: Title back to the default (file name) and Date off.</summary>
    private void ResetFileRow(int rowIndex)
    {
        var row = _fileGrid.Rows[rowIndex];
        string file = row.Cells[_colFile.Index].Value?.ToString() ?? "";
        var f = _files.FirstOrDefault(x => string.Equals(x.FileName, file, StringComparison.OrdinalIgnoreCase));
        string defaultTitle = f != null ? Core.DefaultTitle(f) : "";
        _updating = true;
        row.Cells[_colTitle.Index].Value = defaultTitle;
        row.Cells[_colDate.Index].Value = false;
        _updating = false;
        Log("  Reset " + file + " -> " + defaultTitle);
    }

    /// <summary>
    /// Display the full page title (base name + optional "yyyy-MM" prefix) in the Title column.
    /// The cell's value stays the base title, so the run pipeline can derive the final name
    /// without double-prefixing the date.
    /// </summary>
    private void FileGrid_CellPaint(object? sender, DataGridViewCellPaintingEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex != _colTitle.Index) return;
        if (e.RowIndex == _editingTitleRow) return; // let the edit control show the base title
        var row = _fileGrid.Rows[e.RowIndex];
        string file = row.Cells[_colFile.Index].Value?.ToString() ?? "";
        var f = _files.FirstOrDefault(x => string.Equals(x.FileName, file, StringComparison.OrdinalIgnoreCase));
        string baseTitle = row.Cells[_colTitle.Index].Value?.ToString() ?? "";
        bool date = (bool)(row.Cells[_colDate.Index].Value ?? false);
        string full = f != null ? Core.PageTitle(f, baseTitle, date) : baseTitle;
        var cell = row.Cells[e.ColumnIndex];
        bool isSelected = row.Selected || _fileGrid.SelectedCells.Contains(cell) || _fileGrid.CurrentCell == cell;
        e.PaintBackground(e.CellBounds, isSelected);
        var font = row.DefaultCellStyle.Font ?? _fileGrid.Font;
        using (var br = new SolidBrush(row.DefaultCellStyle.ForeColor))
        {
            var sz = e.Graphics.MeasureString(full, font);
            e.Graphics.DrawString(full, font, br,
                e.CellBounds.Left + 4,
                e.CellBounds.Top + Math.Max(0f, (e.CellBounds.Height - sz.Height) / 2f));
        }
        e.Handled = true;
    }

    private void FileGrid_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
    {
        if (e.ColumnIndex == _colTitle.Index) _editingTitleRow = e.RowIndex;
    }

    private void FileGrid_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.ColumnIndex == _colTitle.Index) _editingTitleRow = -1;
    }

    /// <summary>Repaint the Title preview when its inputs (base title or the Date flag) change.</summary>
    private void FileGrid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        if (e.ColumnIndex == _colDate.Index || e.ColumnIndex == _colTitle.Index)
            _fileGrid.InvalidateCell(_colTitle.Index, e.RowIndex);
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

    /// <summary>Per-file run settings (resolved against the institution mapping).</summary>
    private sealed class RunRow
    {
        public string FileName = "";
        public string Institution = "";
        public bool Enabled = true;
        public TargetOption? Target;
        public string Notebook = "";
        public string Title = "";
        public bool Date = false;
    }

    /// <summary>
    /// Read run settings from both grids: per-file (enabled, title) from the bottom
    /// grid, and the institution's (notebook, target) from the top grid. Marshals a
    /// blocking read onto the UI thread and returns plain data for the worker thread.
    /// </summary>
    private List<RunRow> SnapshotRunRows()
    {
        var rows = new List<RunRow>();
        try
        {
            if (_log == null || !_log.IsHandleCreated) return rows;
            Action fill = () =>
            {
                var instMap = new Dictionary<string, (string Nb, TargetOption? Tgt)>(StringComparer.OrdinalIgnoreCase);
                foreach (DataGridViewRow row in _mapGrid.Rows)
                {
                    string inst = row.Cells[_mapColInst.Index].Value?.ToString() ?? "";
                    string nb = row.Cells[_mapColNotebook.Index].Value?.ToString() ?? "";
                    var tgt = ResolveTarget(row.Cells[_mapColTarget.Index].Value);
                    if (inst.Length > 0) instMap[inst] = (nb, tgt);
                }
                foreach (DataGridViewRow row in _fileGrid.Rows)
                {
                    string file = row.Cells[_colFile.Index].Value?.ToString() ?? "";
                    string inst = row.Cells[_colInst.Index].Value?.ToString() ?? "";
                    string nb = "";
                    TargetOption? tgt = null;
                    if (instMap.TryGetValue(inst, out var mv)) { nb = mv.Nb; tgt = mv.Tgt; }
                    rows.Add(new RunRow
                    {
                        FileName = file,
                        Institution = inst,
                        Enabled = (bool)(row.Cells[_colEnabled.Index].Value ?? true),
                        Target = tgt,
                        Notebook = nb,
                        Title = row.Cells[_colTitle.Index].Value?.ToString() ?? "",
                        Date = (bool)(row.Cells[_colDate.Index].Value ?? false),
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

        int made = 0, skipped = 0, noTarget = 0, disabled = 0, dup = 0;
        int dpi = Core.GetDpi(_env);
        int maxPages = Core.GetMaxPages(_env);
        var problems = new List<string>();
        var results = new List<(string File, string Status, bool Uncheck)>();
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
                results.Add((f.FileName, "disabled", false));
                continue;
            }
            if (rr.Target == null)
            {
                Log($"  {f.FileName}: NO TARGET SET — skipping.", error: true);
                noTarget++;
                skipped++;
                problems.Add($"{f.FileName} — no target set (pick a Notebook + Target in the top grid).");
                results.Add((f.FileName, "no target", false));
                continue;
            }

            string pdf = Path.Combine(_outDir, f.Category, f.FileName);
            if (!File.Exists(pdf))
            {
                Log($"  ! {f.FileName}: PDF not found at {pdf} — skipping.", error: true);
                skipped++;
                problems.Add($"{f.FileName} — PDF not found at {pdf}.");
                results.Add((f.FileName, "no PDF", false));
                continue;
            }

            int year = f.Statements.Select(s => s.StatementDate?.Year ?? 0).Max();
            string pageName = Core.PageTitle(f, rr.Title, rr.Date);
            Log($"  {f.FileName} -> [{rr.Notebook}] {rr.Target} / {pageName}");

            if (dryRun)
            {
                made++;
                results.Add((f.FileName, "would create", false));
                continue;
            }

            try
            {
                var accRows = Core.BuildAccountRows(f);
                var entry = new MapEntry { Kind = rr.Target.Kind, Name = rr.Target.Name };
                var sec = ResolveSection(onenote!, ref h!, entry, year, rr.Notebook, rr.Title, rr.Target.ParentName);
                Log($"    -> [{sec.Notebook}] {sec.Name} / {pageName}");

                // Don't create a page whose title already exists in the target section.
                string existing = onenote!.GetPageNames(sec.ObjectID)
                    .FirstOrDefault(n => string.Equals(n, pageName, StringComparison.Ordinal)) ?? "";
                if (existing.Length > 0)
                {
                    Log($"    page '{pageName}' already exists in [{sec.Notebook}] {sec.Name} — skipping (no duplicate).");
                    dup++;
                    problems.Add($"{f.FileName} — page '{pageName}' already exists in [{sec.Notebook}] {sec.Name} (duplicate).");
                    results.Add((f.FileName, "duplicate", false));
                    continue;
                }

                string pageId = onenote!.CreateNote(sec.ObjectID, pageName);
                var rasters = PdfTools.Rasterize(pdf, dpi, maxPages);
                onenote!.CommitPageFull(pageId, pageName, accRows, pdf, rasters, dpi);
                made++;
                results.Add((f.FileName, "done", true));   // success -> uncheck
                Log($"    created + content set ({rasters.Count} image(s), {f.Statements.Count} account(s)).");
            }
            catch (Exception e)
            {
                skipped++;
                problems.Add($"{f.FileName} — failed: {e.Message}");
                results.Add((f.FileName, "failed", false));
                Log($"    ! {f.FileName} failed: {e.Message}", error: true);
            }
        }

        using (onenote) { }
        OnUi(() => _btnStop.Enabled = false);
        OnUi(() => ApplyResults(results));
        Log($"Done: {made} page(s), {skipped} skipped, {noTarget} without target, {disabled} disabled, {dup} duplicate(s).");
        if (problems.Count > 0)
        {
            Log("");
            Log($"=== NEEDS ATTENTION — {problems.Count} item(s) not created ===", error: true);
            foreach (var p in problems) Log("  - " + p, error: true);
            Log("Those files stayed checked. Fix the issue and run again.", error: true);
        }
    }

    /// <summary>
    /// Apply per-file run results to the bottom grid: set the Status column and
    /// uncheck files that were processed successfully (skipped/failed files stay
    /// checked so they are retried on the next run). Must be called on the UI thread.
    /// </summary>
    private void ApplyResults(List<(string File, string Status, bool Uncheck)> results)
    {
        var dict = new Dictionary<string, (string Status, bool Uncheck)>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in results) dict[r.File] = (r.Status, r.Uncheck);
        foreach (DataGridViewRow row in _fileGrid.Rows)
        {
            string file = row.Cells[_colFile.Index].Value?.ToString() ?? "";
            if (!dict.TryGetValue(file, out var su)) continue;
            row.Cells[_colStatus.Index].Value = su.Status;
            if (su.Uncheck) row.Cells[_colEnabled.Index].Value = false;
        }
    }

    private void DoTest()
    {
        Log("TEST: creating a page in OneNote with a sample table + a generated image ...");
        try
        {
            using var onenote = new OneNote();
            EnsureHierarchy(onenote);
            var target = BuildAllOptions().FirstOrDefault(o => o.Kind == "section");
            if (target == null)
            {
                Log("TEST: no section found in OneNote to test against.", error: true);
                return;
            }
            Log($"TEST target section: {target.Name} in {target.Notebook}");
            Log("TEST: ensuring section ...");
            var sec = ResolveSection(onenote, ref _hierarchy!, new MapEntry { Kind = target.Kind, Name = target.Name }, 2026, target.Notebook, "", target.ParentName);
            Log("TEST: creating page ...");
            string pageId = onenote.CreateNote(sec.ObjectID, "TEST " + DateTime.Now.ToString("HHmmss"));
            Log("TEST page id: " + pageId);
            Log("TEST: generating a sample PNG (rasterizer sanity check) ...");
            int w = 640, hgt = 360;
            byte[] px = new byte[w * hgt];
            for (int y = 0; y < hgt; y++)
                for (int x = 0; x < w; x++)
                {
                    int idx = y * w + x;
                    px[idx * 4 + 0] = (byte)(x * 255 / w);
                    px[idx * 4 + 1] = (byte)(y * 255 / hgt);
                    px[idx * 4 + 2] = (byte)(255 - x * 255 / w);
                    px[idx * 4 + 3] = 255;
                }
            byte[] png = PngWriter.Encode(px, w, hgt);
            Log($"TEST generated {png.Length}-byte PNG");
            var sample = new PageXml.AccountRow
            {
                Date = "2026-06-30",
                Name = "Example Checking (…1234)",
                Balance = "$26,372.47",
                Notes = new List<string> { "Direct deposit — ACME PAYROLL $1,450.00", "ACH DEPOSIT - 123456 - 4321 - 654321" },
            };
            var rasters = new List<(string Name, byte[] Png)> { ("test.png", png) };
            Log("TEST: writing page content (table + image) ...");
            onenote.CommitPageFull(pageId, "TEST " + DateTime.Now.ToString("HHmmss"), new List<PageXml.AccountRow> { sample }, null, rasters, 100);
            Log("TEST: page written. Open OneNote and check the page in section '" + sec.Name + "'.", error: true);
        }
        catch (Exception e)
        {
            Log("TEST FAILED: " + e.Message, error: true);
            if (e is System.Runtime.InteropServices.COMException ce)
                Log("  HRESULT 0x" + ce.HResult.ToString("X8"), error: true);
        }
    }

    /// <summary>
    /// Resolve the target section, creating it if needed.
    ///  - "group"  -> use the group as the parent; ensure a section "{Title} {year}" exists in it.
    ///  - "section"-> find/create the section in that notebook (disambiguated by the
    ///                containing group, resolved from <c>parentName</c> to an ObjectID).
    /// </summary>
    private static SectionInfo ResolveSection(OneNote onenote, ref Hierarchy h, MapEntry entry, int year, string notebook, string title, string parentName)
    {
        // Section-group target: ensure "{Title} {year}" section exists inside it.
        if (entry.Kind == "group")
        {
            string notebookId = NotebookId(h, notebook);
            var g = h.Groups.FirstOrDefault(x => x.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase) && SameParent(x.Parent, notebookId))
                  ?? h.Groups.FirstOrDefault(x => x.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase))
                  ?? throw new Exception($"Section group \"{entry.Name}\" not found in notebook \"{notebook}\".");
            string baseName = (title.Length == 0) ? g.Name : title;
            string name = $"{baseName} {year}";
            var existing = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, g.ObjectID));
            if (existing != null) return existing;
            try
            {
                onenote.CreateSectionInGroup(name, g);
                h = onenote.GetHierarchy();
            }
            catch (Exception e)
            {
                throw new Exception($"Could not create section \"{name}\" in group \"{entry.Name}\" ({e.Message}). " +
                                   "If OneNote is blocked from creating it, create it manually and re-run.");
            }
            var created = h.Sections.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, g.ObjectID));
            if (created != null) return created;
            throw new Exception($"Section \"{name}\" was not created (may be blocked by a lock). Create it manually in OneNote and re-run.");
        }

        // Section target — disambiguate by containing group.
        // Resolve the (persisted) group NAME to an ObjectID for comparison.
        string parentID = "";
        if (!string.IsNullOrEmpty(parentName))
            parentID = h.Groups.FirstOrDefault(x => x.Name.Equals(parentName, StringComparison.OrdinalIgnoreCase)
                && (string.IsNullOrEmpty(notebook) || x.Notebook.Equals(notebook, StringComparison.OrdinalIgnoreCase)))?.ObjectID ?? "";

        var sec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, parentID));
        if (sec != null) return sec;

        try
        {
            if (string.IsNullOrEmpty(parentID))
                onenote.CreateSection(entry.Name, string.IsNullOrEmpty(notebook) ? secNotebookFallback(h) : notebook);
            else
            {
                var g = h.Groups.FirstOrDefault(x => SameParent(x.ObjectID, parentID))
                      ?? throw new Exception($"Section group \"{parentName}\" not found (needed to create section \"{entry.Name}\").");
                onenote.CreateSectionInGroup(entry.Name, g);
            }
            h = onenote.GetHierarchy();
        }
        catch (Exception e)
        {
            throw new Exception($"Could not create section \"{entry.Name}\" ({e.Message}). Create it manually in OneNote and re-run.");
        }

        var createdSec = h.Sections.FirstOrDefault(s => s.Name.Equals(entry.Name, StringComparison.OrdinalIgnoreCase) && SameParent(s.Parent, parentID));
        if (createdSec != null) return createdSec;
        throw new Exception($"Section \"{entry.Name}\" was not created (may be blocked by a lock). Create it manually in OneNote and re-run.");
    }

    private static bool SameParent(string a, string b) => string.Equals(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase);

    private static string NotebookId(Hierarchy h, string notebookName)
    {
        var nb = h.Notebooks.FirstOrDefault(n => n.Name.Equals(notebookName, StringComparison.OrdinalIgnoreCase));
        return nb?.ObjectID ?? "";
    }

    private static string secNotebookFallback(Hierarchy h)
    {
        var nb = h.Notebooks.FirstOrDefault(n => !n.Name.StartsWith("OneNote_RecycleBin", StringComparison.OrdinalIgnoreCase));
        return nb?.Name ?? "StatementsAll";
    }

    /// <summary>Toggle all file rows on/off (on if any is off, off if all on).</summary>
    private void ToggleAll()
    {
        if (_fileGrid.Rows.Count == 0) return;
        bool allOn = true;
        foreach (DataGridViewRow r in _fileGrid.Rows)
            if (!(r.Cells[_colEnabled.Index].Value is bool b && b)) { allOn = false; break; }
        _updating = true;
        try
        {
            foreach (DataGridViewRow r in _fileGrid.Rows)
                r.Cells[_colEnabled.Index].Value = !allOn;
            Log($"All file rows set to {(allOn ? "OFF" : "ON")} ({_fileGrid.Rows.Count} rows).");
        }
        finally { _updating = false; }
    }

    /// <summary>Persist the grids' state to onemap.json: institution mapping (top) + per-file settings (bottom).</summary>
    private void SaveMappingFromGrid(bool silent = false)
    {
        var map = new OneMap();

        // Per-institution mapping from the TOP grid.
        foreach (DataGridViewRow row in _mapGrid.Rows)
        {
            string inst = row.Cells[_mapColInst.Index].Value?.ToString() ?? "";
            if (inst.Length == 0) continue;
            var opt = ResolveTarget(row.Cells[_mapColTarget.Index].Value);
            if (opt != null)
            {
                string nb = row.Cells[_mapColNotebook.Index].Value?.ToString() ?? "";
                map.Institutions[inst] = new MapEntry
                {
                    Kind = opt.Kind,
                    Name = opt.Name,
                    Notebook = string.IsNullOrEmpty(nb) ? opt.Notebook : nb,
                    Parent = opt.ParentName,
                };
            }
        }

        // Per-file settings from the BOTTOM grid.
        foreach (DataGridViewRow row in _fileGrid.Rows)
        {
            string file = row.Cells[_colFile.Index].Value?.ToString() ?? "";
            if (file.Length == 0) continue;
            map.Files[file] = new FileSetting
            {
                Enabled = (bool)(row.Cells[_colEnabled.Index].Value ?? true),
                Title = row.Cells[_colTitle.Index].Value?.ToString() ?? "",
                Date = (bool)(row.Cells[_colDate.Index].Value ?? false),
            };
        }

        try
        {
            string mapPath = Core.MapPath(_root);
            Core.SaveMap(mapPath, map);
            if (!silent)
                Log($"Saved {map.Institutions.Count} institution mapping(s) and {map.Files.Count} file setting(s) to {mapPath}");
        }
        catch (Exception e)
        {
            Log("Save mapping failed: " + e.Message, error: true);
        }
    }

    // ------------------------------------------------------------------
    // Logging
    // ------------------------------------------------------------------

    /// <summary>Thread-safe log: marshals to the UI thread, buffering if it isn't ready yet.</summary>
    private void Log(string message, bool error = false)
    {
        var s = (error ? "[!] " : "") + message;
        try
        {
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
        catch
        {
            _pendingLog.Add(s + "|");
        }
    }

    /// <summary>Drain any log lines captured before the UI thread / handle was ready.</summary>
    private void FlushPendingLog()
    {
        if (_pendingLog.Count == 0) return;
        var snapshot = _pendingLog.ToList();
        _pendingLog.Clear();
        foreach (var s in snapshot)
            AppendLogLine(s.TrimEnd('|'));
    }

    // ------------------------------------------------------------------
    // Console redirect (keeps Console.* writes from being lost in a WinExe)
    // ------------------------------------------------------------------

    private sealed class ConsoleToLog : StreamWriter
    {
        private readonly Action<string> _sink;
        public ConsoleToLog(Action<string> sink) : base(new MemoryStream(), Encoding.UTF8) { _sink = sink; AutoFlush = true; }
        public override void Write(string? value) { if (!string.IsNullOrEmpty(value)) _sink(value); }
        public override void WriteLine(string? value) { if (value != null) _sink(value); }
    }
}
