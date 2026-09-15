using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.OneNote;

class P {
    [System.STAThread]
    static void Main() {
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE") { UseShellExecute = true }); } catch {}
        System.Threading.Thread.Sleep(3000);
        var app = new ApplicationClass();
        string xml;
        app.GetHierarchy("", HierarchyScope.hsSections, out xml);

        // correctly extract (name, ID) for each <one:Section ... ID="...">
        var secs = new List<Tuple<string,string>>();
        foreach (Match m in Regex.Matches(xml, "<one:Section\\s+name=\"([^\"]*)\"\\s+ID=\"([^\"]*)\"")) {
            string name = m.Groups[1].Value, id = m.Groups[2].Value;
            if (!secs.Any(t => t.Item2 == id)) secs.Add(Tuple.Create(name, id));
        }
        Console.WriteLine("Collected " + secs.Count + " real section(s).");
        foreach (var s in secs.Take(5)) Console.WriteLine("  " + s.Item1 + "  " + s.Item2);
        if (secs.Count == 0) { Console.WriteLine("none found; dumping 1500 chars:"); Console.WriteLine(xml.Substring(0, Math.Min(1500, xml.Length))); return; }

        // Try CreateNewPage on the first several real sections (some may be OneDrive 'areAllPagesAvailable=false')
        for (int i = 0; i < Math.Min(6, secs.Count); i++) {
            TryCreate(app, secs[i].Item1, secs[i].Item2);
        }
    }

    static void TryCreate(ApplicationClass app, string name, string id) {
        try {
            app.CreateNewPage(id, out string pid);
            Console.WriteLine("OK   page=" + pid + "  in section '" + name + "'");
        }
        catch (Exception e) {
            Console.WriteLine("FAIL 0x" + e.HResult.ToString("X8") + "  section '" + name + "'");
        }
    }
}
