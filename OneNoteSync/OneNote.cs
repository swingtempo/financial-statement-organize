using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OneNoteSync;

public sealed record SectionInfo(string Guid, string Name, string Notebook);
public sealed record GroupInfo(string Guid, string Name, string Notebook);
public sealed record Hierarchy(IList<SectionInfo> Sections, IList<GroupInfo> Groups);

/// <summary>
/// Thin wrapper over the OneNote desktop COM API (OneNote.Application).
/// Uses dynamic dispatch so no interop package is required — OneNote
/// (Microsoft 365 / 2016 desktop) must be installed.
/// </summary>
public sealed class OneNote : IDisposable
{
    private dynamic? _app;

    public OneNote()
    {
        var type = Type.GetTypeFromProgID("OneNote.Application");
        if (type == null)
            throw new InvalidOperationException(
                "OneNote COM API not found. Install the OneNote desktop app " +
                "(Microsoft 365 / OneNote 2016) — the Windows 10 (UWP) OneNote does not expose this API.");
        _app = Activator.CreateInstance(type)!;
    }

    /// <summary>Get the notebook hierarchy (sections + section groups).</summary>
    public Hierarchy GetHierarchy()
    {
        // GetHierarchyScope: 0 = entireNotebooks, 2 = visibleNotebooks, 1 = activeNotebook.
        // entireNotebooks can fail (e.g. a closed notebook); fall back to narrower scopes.
        string? xml = null;
        int lastError = 0;
        foreach (int scope in new[] { 0, 2, 1 })
        {
            try { xml = (string)_app!.GetHierarchy((uint)scope); break; }
            catch (COMException e) { lastError = e.HResult; }
        }
        if (xml == null)
            throw new Exception($"Could not read the OneNote hierarchy (HRESULT 0x{lastError:X8}). " +
                              "Make sure the OneNote desktop app (Microsoft 365 / 2016) is installed " +
                              "and at least one notebook is open, then retry.");
        var doc = XDocument.Parse(xml);
        var sections = new List<SectionInfo>();
        var groups = new List<GroupInfo>();
        foreach (var nb in doc.Root?.Elements("notebook") ?? Enumerable.Empty<XElement>())
        {
            string nbName = (string?)nb.Attribute("name") ?? "?";
            foreach (var s in nb.Descendants("section"))
                sections.Add(new SectionInfo((string)s.Attribute("guid")!, (string?)s.Attribute("name") ?? "?", nbName));
            foreach (var g in nb.Descendants("sectionGroup"))
                groups.Add(new GroupInfo((string)g.Attribute("guid")!, (string?)g.Attribute("name") ?? "?", nbName));
        }
        return new Hierarchy(sections, groups);
    }

    /// <summary>Create a new section (created in the default location).</summary>
    public void CreateSection(string name) => _app!.UpdateSection($"<section name=\"{EscapeXml(name)}\"></section>");

    /// <summary>Move an existing section into a section group.</summary>
    public void MoveSectionToGroup(string sectionGuid, string groupGuid)
        => _app!.MoveSection(sectionGuid, groupGuid, 0);

    /// <summary>Create a new page; returns the page URI.</summary>
    public string CreateNote(string sectionGuid, string pageName)
        => (string)_app!.CreateNewNote(sectionGuid, pageName);

    /// <summary>Replace the page body with the given HTML (images may use base64 data URIs).</summary>
    public void UpdatePageHtml(string pageUri, string html)
    {
        // ContentScope.page = 0, display option 0 (default)
        string content = (string)_app!.GetContent(pageUri, 0u, 0u);
        var m = Regex.Match(content, @"(<one:Xml>).*?(</one:Xml>)", RegexOptions.Singleline);
        if (m.Success)
        {
            int end = m.Groups[1].Index + m.Groups[1].Length;
            content = content[..end] + "<![CDATA[" + html + "]]>" + content[end..];
        }
        else
        {
            // Fallback: minimal page document.
            content = "<one:Page><one:Source><one:Html><one:Xml><![CDATA[" + html +
                    "]]></one:Xml></one:Html></one:Source></one:Page>";
        }
        _app!.UpdatePage(pageUri, content);
    }

    /// <summary>Attach one or more files (e.g. a PDF) to the page.</summary>
    public void AddFilesToPage(string pageUri, params string[] files) => _app!.AddFilesToPage(pageUri, files);

    public void Dispose()
    {
        _app = null;
    }

    private static string EscapeXml(string s)
        => s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");
}
