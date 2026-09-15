using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace OneNoteSync
{
    public sealed class SectionInfo
    {
        public string Name { get; set; } = "";
        public string ObjectID { get; set; } = "";
        public string Notebook { get; set; } = "";
        public string Parent { get; set; } = "";   // containing section-group name ("" if a direct notebook child)
    }

    public sealed class GroupInfo
    {
        public string Name { get; set; } = "";
        public string ObjectID { get; set; } = "";
        public string Notebook { get; set; } = "";
        public string Parent { get; set; } = "";   // containing notebook name
    }

    public sealed class NotebookInfo
    {
        public string Name { get; set; } = "";
        public string ObjectID { get; set; } = "";
    }

    public sealed class Hierarchy
    {
        public List<NotebookInfo> Notebooks { get; } = new();
        public List<SectionInfo> Sections { get; } = new();
        public List<GroupInfo> Groups { get; } = new();
    }

    /// <summary>
    /// High-level OneNote operations: hierarchy parsing, section creation,
    /// page creation, and page-content setting (title via CreateNewPage, body via
    /// the proven UpdatePageContent path with summary + PDF + page images).
    /// </summary>
    public sealed class OneNote : IDisposable
    {
        readonly OneNoteApp _app;
        public string Launch => _app.Launch;

        public OneNote()
        {
            _app = new OneNoteApp();
        }

        // ---------------------------------------------------------------- hierarchy

        public Hierarchy GetHierarchy()
        {
            string xml = _app.GetHierarchy();
            return ParseHierarchy(xml);
        }

        static Hierarchy ParseHierarchy(string xml)
        {
            var h = new Hierarchy();
            if (string.IsNullOrWhiteSpace(xml)) return h;
            try
            {
                var root = XDocument.Parse(xml).Root;
                if (root == null) return h;
                foreach (var nb in root.DescendantsAndSelf().Where(e => e.Name.LocalName == "Notebook"))
                {
                    string nbName = (string)nb.Attribute("name") ?? "";
                    string nbId = (string)nb.Attribute("ID") ?? (string)nb.Attribute("objectID") ?? "";
                    h.Notebooks.Add(new NotebookInfo { Name = nbName, ObjectID = nbId });

                    // Sections that are DIRECT children of the notebook (no containing group).
                    foreach (var sec in nb.Elements().Where(e => e.Name.LocalName == "Section"))
                    {
                        h.Sections.Add(new SectionInfo
                        {
                            Name = (string)sec.Attribute("name") ?? "",
                            ObjectID = (string)sec.Attribute("ID") ?? (string)sec.Attribute("objectID") ?? "",
                            Notebook = nbName
                        });
                    }

                    // Section groups (parent = this notebook).
                    foreach (var grp in nb.Elements().Where(e => e.Name.LocalName == "SectionGroup"))
                    {
                        string gName = (string)grp.Attribute("name") ?? "";
                        var g = new GroupInfo
                        {
                            Name = gName,
                            ObjectID = (string)grp.Attribute("ID") ?? (string)grp.Attribute("objectID") ?? "",
                            Notebook = nbName,
                            Parent = nbId
                        };
                        h.Groups.Add(g);

                        // Sections inside the group (parent = this group).
                        foreach (var sec in grp.Elements().Where(e => e.Name.LocalName == "Section"))
                        {
                            h.Sections.Add(new SectionInfo
                            {
                                Name = (string)sec.Attribute("name") ?? "",
                                ObjectID = (string)sec.Attribute("ID") ?? (string)sec.Attribute("objectID") ?? "",
                                Notebook = nbName,
                                Parent = g.ObjectID
                            });
                        }
                    }
                }
            }
            catch { /* best-effort; return what we have */ }
            return h;
        }

        // ---------------------------------------------------------------- sections

        /// <summary>Add a new section to a notebook (best-effort; may fail with 0x80042004).</summary>
        public void CreateSection(string name, string notebook)
        {
            string x =
                "<one:Notebook name=\"" + X(notebook) + "\">" +
                "<one:Section name=\"" + X(name) + "\"></one:Section>" +
                "</one:Notebook>";
            _app.UpdateHierarchy(x);
        }

        /// <summary>Add a new section inside a section group (best-effort; often blocked by 0x80042004).</summary>
        public void CreateSectionInGroup(string name, GroupInfo group)
        {
            string pgid = "{" + Guid.NewGuid().ToString("N").ToUpperInvariant() + "}{A0}{B0}";
            string x =
                "<one:Notebook name=\"" + X(group.Notebook) + "\" ID=\"" + X(group.Parent) + "\">" +
                "<one:SectionGroup name=\"" + X(group.Name) + "\" ID=\"" + X(group.ObjectID) + "\">" +
                "<one:Section name=\"" + X(name) + "\" ID=\"" + X(pgid) + "\"></one:Section>" +
                "</one:SectionGroup></one:Notebook>";
            _app.UpdateHierarchy(x);
        }

        // ---------------------------------------------------------------- pages

        /// <summary>Create a page in a section (makes the section current, then CreateNewPage).</summary>
        public string CreateNote(string sectionObjectID, string pageName)
        {
            // CreateNewPage(sectionId, out pageId) — creates the page in the given section.
            string id = _app.CreatePage(sectionObjectID);
            if (string.IsNullOrEmpty(id))
                throw new Exception("CreateNewPage returned no page id (LastError: " + (_app.LastError?.Message ?? "n/a") + ")");
            return id;
        }

        /// <summary>Set a page's body to summary text only.</summary>
        public void CommitPage(string pageId, string pageName, List<string> lines)
        {
            SetPageBody(pageId, string.Join("\n", lines), null, null, null, 150);
        }

        /// <summary>Set a page's body to summary + PDF attachment + rendered page images.</summary>
        public void CommitPageFull(string pageId, string summary, string pdfPath,
                                  List<(string Name, byte[] Png)> rasters, int dpi)
        {
            SetPageBody(pageId, summary, pdfPath,
                       string.IsNullOrEmpty(pdfPath) ? null : System.IO.Path.GetFileName(pdfPath),
                       rasters, dpi);
        }

        /// <summary>Set a page's body from pre-encoded binary (used by the test page).</summary>
        public (bool imgOk, bool fileOk, string note) TryCommitPageWithBinary(
            string pageId, string pageName, List<string> lines,
            string imgB64, int w, int h, string fileB64, string fileName)
        {
            string summary = string.Join("\n", lines);
            // Rebuild an in-memory raster list from the supplied base64 (single image).
            var rasters = new List<(string Name, byte[] Png)>();
            if (!string.IsNullOrEmpty(imgB64))
                rasters.Add(("test", Convert.FromBase64String(imgB64)));

            string pdfPath = null;
            if (!string.IsNullOrEmpty(fileB64))
            {
                // Write the PDF bytes to a temp file so pathSource can point at it.
                string tmp = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                    "stmt_" + Guid.NewGuid().ToString("N") + "_" + (fileName ?? "document.pdf"));
                System.IO.File.WriteAllBytes(tmp, Convert.FromBase64String(fileB64));
                pdfPath = tmp;
            }

            try
            {
                SetPageBody(pageId, summary, pdfPath,
                           string.IsNullOrEmpty(pdfPath) ? null : System.IO.Path.GetFileName(pdfPath),
                           rasters, 150);
                bool imgOk = rasters.Count > 0;
                bool fileOk = !string.IsNullOrEmpty(pdfPath);
                return (imgOk, fileOk, "Committed via UpdatePageContent (pathSource + inline base64 image).");
            }
            catch (Exception e)
            {
                return (false, false, "UpdatePageContent failed: " + e.Message);
            }
        }

        // ---------------------------------------------------------------- internals

        void SetPageBody(string pageId, string summary, string pdfPath, string pdfName,
                        List<(string Name, byte[] Png)> rasters, int dpi)
        {
            //try { _app.Navigate(pageId); } catch { }
            System.Threading.Thread.Sleep(500);
            string xml = _app.GetPage(pageId, 0); // piBasic
            if (string.IsNullOrEmpty(xml))
                throw new Exception("GetPageContent returned empty for page " + pageId);
            string body = PageXml.BuildBody(summary, pdfPath, pdfName, rasters, dpi);
            int idx = xml.LastIndexOf("</one:Page>");
            if (idx < 0) throw new Exception("No </one:Page> in page XML");
            string newXml = xml.Insert(idx, body);
            _app.UpdatePage(newXml);
        }

        static string X(string s) => (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

        public void Dispose() { _app.Dispose(); }
    }
}
