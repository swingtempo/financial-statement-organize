using System;
using Microsoft.Office.Interop.OneNote;

class P {
    static void Main() {
        try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(@"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE") { UseShellExecute = true }); } catch {}
        System.Threading.Thread.Sleep(3000);
        var app = new ApplicationClass();
        string xml;
        app.GetHierarchy("", HierarchyScope.hsSections, out xml);
        string[] lines = xml.Split('\n');
        for (int i = 0; i < lines.Length; i++) {
            if (lines[i].Contains("name=\"Family Notebook\"")) {
                for (int j = i; j < lines.Length && j < i + 60; j++) Console.WriteLine(lines[j]);
                break;
            }
        }
    }
}
