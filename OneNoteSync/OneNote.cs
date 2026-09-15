using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteSync
{
    /// <summary>
    /// OneNote COM wrapper (net48, GAC PIA) — EARLY-BOUND, exactly like the
    /// working OneNoteXmlViewer and the prior committed version.
    ///
    /// (The earlier reflection-based version failed with "Library not registered"
    ///  because GetType().GetTypeInfo().InvokeMember on the PIA object can't
    ///  resolve the type library. Early-bound IApplication calls work.)
    ///
    /// On boot: launch OneNote if it isn't running (per request), then connect.
    /// </summary>
    public sealed class OneNoteApp : IDisposable
    {
        IApplication _app;
        bool _owned;
        public string Launch { get; private set; } = "";
        public Exception LastError { get; private set; }

        public OneNoteApp()
        {
            EnsureOneNoteRunning();

            // The connect is reliable; a short retry rides out a slow OneNote start.
            int attempt = 0;
            while (true)
            {
                attempt++;
                try
                {
                    _app = new ApplicationClass();
                    _owned = true;
                    Launch += " -> connected (attempt " + attempt + ")";
                    return;
                }
                catch (Exception e)
                {
                    Launch += " [attempt " + attempt + " fail: " + e.Message + "]";
                    if (attempt >= 10)
                        throw new Exception("Cannot get OneNote COM app. Tried: " + Launch);
                    Thread.Sleep(1500);
                }
            }
        }

        /// <summary>Launch OneNote if it isn't already running (auto-boot on start).</summary>
        public static void EnsureOneNoteRunning()
        {
            var p = Process.GetProcessesByName("ONENOTE");
            if (p.Length > 0) { p[0].Dispose(); return; }
            string exe = @"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE";
            if (File.Exists(exe))
            {
                try { Process.Start(exe); } catch { }
                Thread.Sleep(5000);
            }
        }

        // ---------------------------------------------------------------- early-bound PIA

        public void UpdatePage(string xml) => _app.UpdatePageContent(xml);

        public void Navigate(string id) => _app.NavigateTo(id);

        public string GetPage(string id, int info)
        {
            string xml = null;
            _app.GetPageContent(id, out xml, (PageInfo)info);
            return xml ?? "";
        }

        public string GetHierarchy()
        {
            string xml = null;
            _app.GetHierarchy("", HierarchyScope.hsSections, out xml);
            return xml ?? "";
        }

        // CreateNewPage(String bstrSectionID, out String pbstrPageID) — creates a page
        // in the given section (the first arg is the SECTION id, not a page name).
        public string CreatePage(string sectionGuid)
        {
            string id = null;
            _app.CreateNewPage(sectionGuid, out id);
            return id ?? "";
        }

        public void UpdateHierarchy(string xml) => _app.UpdateHierarchy(xml);

        public void Dispose()
        {
            if (_owned && _app != null)
            {
                try { Marshal.ReleaseComObject(_app); } catch { }
            }
            _app = null;
        }
    }
}
