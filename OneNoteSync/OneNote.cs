using Microsoft.Office.Interop.OneNote;
using System.Text;
using System.Xml.Linq;

namespace OneNoteSync;

public sealed record SectionInfo(string Guid, string Name, string Notebook, string? Path = null);
public sealed record GroupInfo(string Guid, string Name, string Notebook);
public sealed record Hierarchy(IList<SectionInfo> Sections, IList<GroupInfo> Groups);

/// <summary>
/// Thin wrapper over the OneNote desktop COM API using the genuine PIA
/// (Microsoft.Office.Interop.OneNote). Requires the OneNote desktop app
/// (Microsoft 365 / OneNote 2016) — the Windows 10 (UWP) OneNote does not
/// expose this API.
/// </summary>
public sealed class OneNote : IDisposable
{
    private readonly IApplication _app;
    public object RawAppObj() => _app;
    public System.Type RawAppType() => _app.GetType();

    public OneNote()
    {
        // Instantiate via the PIA CoClass (CLSID D7FAC39E…) — the object that
        // actually implements the modern IApplication (verified to work).
        _app = new ApplicationClass();
    }

    /// <summary>Return the raw hierarchy XML (for debugging).</summary>
    public string GetRawHierarchy()
    {
        string? xml = null;
        _app.GetHierarchy("", HierarchyScope.hsSections, out xml);
        return xml ?? "";
    }

    /// <summary>Get the notebook hierarchy (sections + section groups), with section IDs.</summary>
    public Hierarchy GetHierarchy()
    {
        string? xml = null;
        Exception? last = null;
        for (int attempt = 1; attempt <= 5; attempt++)
        {
            try
            {
                string? outXml = null;
                _app.GetHierarchy("", HierarchyScope.hsSections, out outXml);
                xml = outXml;
                if (!string.IsNullOrEmpty(xml)) break;
            }
            catch (Exception e)
            {
                last = e;
            }
            if (string.IsNullOrEmpty(xml) && attempt < 5)
            {
                Console.Error.WriteLine("  GetHierarchy failed (" + last!.Message + "); retrying in 3s (" + attempt + "/4)...");
                Thread.Sleep(3000);
            }
        }
        if (string.IsNullOrEmpty(xml))
            throw new Exception($"Could not read the OneNote hierarchy after retries (last error: {last?.Message}). " +
                              "Make sure the OneNote desktop app is fully loaded and at least one notebook is open, then retry.");
        return Parse(xml!);
    }

    private static Hierarchy Parse(string xml)
    {
        var doc = XDocument.Parse(xml);
        var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
        var sections = new List<SectionInfo>();
        var groups = new List<GroupInfo>();
        foreach (var nb in doc.Root!.Descendants(ns + "Notebook"))
        {
            string nbName = (string?)nb.Attribute("name") ?? "?";
            foreach (var s in nb.Descendants(ns + "Section"))
            {
                string id = (string?)s.Attribute("ID") ?? (string?)s.Attribute("objectID") ?? "";
                sections.Add(new SectionInfo(id, (string?)s.Attribute("name") ?? "?", nbName,
                                           (string?)s.Attribute("path")));
            }
            foreach (var g in nb.Descendants(ns + "SectionGroup"))
            {
                string id = (string?)g.Attribute("ID") ?? (string?)g.Attribute("objectID") ?? "";
                groups.Add(new GroupInfo(id, (string?)g.Attribute("name") ?? "?", nbName));
            }
        }
        return new Hierarchy(sections, groups);
    }

    /// <summary>
    /// Open (or, with create, create) an object in the hierarchy and return its
    /// object ID. For creating a section directly inside a section group, pass the
    /// group's object ID as <paramref name="relativeToObjectID"/>.
    /// </summary>
    public string OpenHierarchy(string path, string relativeToObjectID, CreateFileType createIfNotExist)
    {
        string? id = null;
        _app.OpenHierarchy(path, relativeToObjectID, out id, createIfNotExist);
        return id!;
    }

    /// <summary>Create a new section (in the default location) and return its ID.</summary>
    public string CreateSection(string name) => OpenHierarchy(name, "", CreateFileType.cftSection);

    /// <summary>Create a new section directly inside a section group; returns the section ID.</summary>
    public string CreateSectionInGroup(string sectionName, string groupObjectID)
        => OpenHierarchy(sectionName, groupObjectID, CreateFileType.cftSection);

    /// <summary>Open an existing section group and return its object ID.</summary>
    public string OpenGroup(string groupName) => OpenHierarchy(groupName, "", CreateFileType.cftNone);

    /// <summary>
    /// Create a new blank page in the given section (by section ID). The page
    /// title is supplied by the first &lt;h1&gt; in the content you write to it.
    /// </summary>
    public string CreateNote(string sectionId, NewPageStyle style = NewPageStyle.npsBlankPageNoTitle)
    {
        string? id = null;
        Console.Error.WriteLine($"    [OneNote] CreateNewPage in section {sectionId}");
        _app.CreateNewPage(sectionId, out id, style);
        Console.Error.WriteLine($"    [OneNote] CreateNewPage OK -> {id}");
        return id!;
    }

    /// <summary>Get a page's XML.</summary>
    public string GetPageContent(string pageId)
    {
        string? xml = null;
        Console.Error.WriteLine($"    [OneNote] GetPageContent: {pageId}");
        _app.GetPageContent(pageId, out xml, PageInfo.piBasic);
        Console.Error.WriteLine($"    [OneNote] GetPageContent OK: {(xml == null ? "null" : xml.Length + " chars")}");
        return xml!;
    }

    /// <summary>
    /// Commit changes to a page (full page XML). The PIA early-bound
    /// UpdatePageContent overloads are misaligned on this machine (all return
    /// 0x80042001), so this routes the write through the late-bound PowerShell
    /// helper (OneNoteUpdate.ps1), which dispatches the 1-arg overload correctly.
    /// </summary>
    public void UpdatePageContent(string pageId, string pageXml)
    {
        Console.Error.WriteLine($"    [OneNote] UpdatePageContent: xml length {pageXml.Length}");
        string dir = AppDomain.CurrentDomain.BaseDirectory;
        string helper = null;
        string[] candidates = new[] { dir, Path.Combine(dir, ".."), Path.Combine(dir, "..", ".."), Path.Combine(dir, "..", "..", "..") };
        foreach (var c in candidates) { var p = Path.Combine(c, "OneNoteUpdate.ps1"); if (File.Exists(p)) { helper = p; break; } }
        if (helper == null) throw new System.Exception("OneNoteUpdate.ps1 not found near " + dir);
        string tmp = Path.Combine(Path.GetTempPath(), "onenote-" + Guid.NewGuid().ToString("N") + ".xml");
        File.WriteAllText(tmp, pageXml, System.Text.Encoding.UTF8);
        try
        {
            var psi = new System.Diagnostics.ProcessStartInfo(
                "powershell.exe",
                $"-NoProfile -ExecutionPolicy Bypass -File \"{helper}\" -PageId \"{pageId}\" -XmlFile \"{tmp}\"")
            { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
            using var p = System.Diagnostics.Process.Start(psi)!;
            string outp = p.StandardOutput.ReadToEnd();
            string err = p.StandardError.ReadToEnd();
            p.WaitForExit(120000);
            int code = p.HasExited ? p.ExitCode : -1;
            if (code == 0 && outp.Contains("UPDATE_OK"))
            {
                Console.Error.WriteLine($"    [OneNote] UpdatePageContent OK (sidecar)");
            }
            else
            {
                throw new System.Exception($"UpdatePageContent sidecar failed (code {code}): {outp.Trim()} {err.Trim()}");
            }
        }
        finally { try { File.Delete(tmp); } catch { } }
    }
    private static string GetHr(System.Exception e)
        => e is System.Runtime.InteropServices.COMException ce ? $"0x{ce.HResult:X8}" : (e is System.Runtime.InteropServices.SEHException se ? $"0x{se.HResult:X8}" : "");

    /// <summary>
    /// Commit a page with a title and a list of text lines (the account
    /// summary). Uses the confirmed-working OneNote 2013 page model:
    /// Page &gt; Title + Outline &gt; OEChildren &gt; OE &gt; T. The write goes
    /// through the late-bound sidecar (see UpdatePageContent).
    /// </summary>
    public void CommitPage(string pageId, string title, IEnumerable<string> lines)
    {
        string tmpl = GetPageContent(pageId);
        int cp = tmpl.IndexOf("</one:Page>", StringComparison.Ordinal);
        if (cp < 0) throw new Exception("Page template has no </one:Page> marker");
        string body = BuildTitle(title) + BuildOutline(lines.ToArray());
        UpdatePageContent(pageId, tmpl.Substring(0, cp) + body + tmpl.Substring(cp));
    }

    /// <summary>
    /// Prototype: commit the working core (title + summary lines) first, then
    /// attempt to append a rasterized image (PDF printout) and/or a file
    /// attachment. The OneNote binary-part mechanism (CallbackID / MediaIndex)
    /// could not be registered through UpdatePageContent in testing, so this
    /// reports back whether the binary parts are actually readable afterwards.
    /// The core (title + text) is always committed regardless of the binary
    /// result.
    /// </summary>
    public (bool ImageOk, bool FileOk, string Note) TryCommitPageWithBinary(
        string pageId, string title, IEnumerable<string> lines,
        string? imageB64, int imageW, int imageH, string? fileB64, string fileName)
    {
        // 1) Always commit the working core first (title + text lines).
        CommitPage(pageId, title, lines);

        bool hasImage = imageB64 != null, hasFile = fileB64 != null;
        if (!hasImage && !hasFile) return (false, false, "no binary requested");

        // 2) Attempt to append the binary elements into the existing outline.
        string imgCb = "{" + Guid.NewGuid().ToString("N").ToUpperInvariant() + "}{1}{B0}";
        string fileCb = "{" + Guid.NewGuid().ToString("N").ToUpperInvariant() + "}{1}{B1}";
        var add = new StringBuilder();
        if (hasImage)
            add.Append("<one:OE><one:Image><one:Size width='" + imageW + "' height='" + imageH +
                      "'/><one:CallbackID callbackID='" + imgCb + "'/></one:Image></one:OE>");
        if (hasFile)
            add.Append("<one:OE><one:FileAttachment><one:Size width='0' height='0'/><one:Meta name='Filename' value='" +
                      EscapeXml(fileName ?? "") + "'/><one:CallbackID callbackID='" + fileCb +
                      "'/></one:FileAttachment></one:OE>");

        bool appendOk = false;
        try
        {
            string cur = GetPageContent(pageId);
            int marker = cur.IndexOf("</one:OEChildren>", StringComparison.Ordinal);
            if (marker >= 0)
            {
                string next = cur.Substring(0, marker) + add.ToString() + cur.Substring(marker);
                UpdatePageContent(pageId, next);
                appendOk = true;
            }
        }
        catch { appendOk = false; }

        // 3) Verify whether the binary parts were actually registered.
        bool imgOk = false, fileOk = false;
        if (hasImage) imgOk = appendOk && ProbeBinary(pageId, imgCb);
        if (hasFile) fileOk = appendOk && ProbeBinary(pageId, fileCb);

        var noteParts = new List<string>();
        if (hasImage) noteParts.Add($"image part {(imgOk ? "registered" : (appendOk ? "NOT registered (shows as a placeholder)" : "append rejected"))}");
        if (hasFile) noteParts.Add($"file part {(fileOk ? "registered" : (appendOk ? "NOT registered (shows as a placeholder)" : "append rejected"))}");
        return (imgOk, fileOk, string.Join("; ", noteParts.ToArray()));
    }

    private bool ProbeBinary(string pageId, string callbackId)
    {
        try
        {
            string b64;
            _app.GetBinaryPageContent(pageId, callbackId, out b64);
            return !string.IsNullOrEmpty(b64);
        }
        catch { return false; }
    }

    private static string BuildTitle(string title)
        => "<one:Title><one:OE><one:T><![CDATA[" + Cdata(title) + "]]></one:T></one:OE></one:Title>";

    private static string BuildOutline(string[] lines)
    {
        var sb = new StringBuilder();
        foreach (var line in lines)
            sb.Append("<one:OE><one:T><![CDATA[" + Cdata(line) + "]]></one:T></one:OE>");
        return "<one:Outline><one:Position x='36' y='86' z='1'/><one:OEChildren>" + sb +
               "</one:OEChildren></one:Outline>";
    }

    /// <summary>Escape a string for use inside a CDATA section.</summary>
    private static string Cdata(string s) => (s ?? "").Replace("]]>", "]]]]><![CDATA[>");

    public void Dispose()
    {
        if (_app != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
    }

    private static string EscapeXml(string s)
        => s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");

    internal static string DescribeHr(Exception? e)
    {
        if (e is System.Runtime.InteropServices.COMException ce) return $"HRESULT 0x{ce.HResult:X8}";
        if (e is System.Runtime.InteropServices.SEHException se) return $"HRESULT 0x{se.HResult:X8}";
        return "";
    }
}
