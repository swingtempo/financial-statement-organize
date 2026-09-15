using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

class P {
    static void Main(string[] args) {
        // Reproduce ParseHierarchy's notebook iteration against the raw dump.
        // We'll feed it XML via file to avoid OneNote.
        string xml = File.ReadAllText(args[0]);
        var root = XElement.Load(new StringReader(xml));
        Console.WriteLine("root = " + root.Name.LocalName);
        var nbs = root.DescendantsAndSelf().Where(e => e.Name.LocalName == "Notebook").ToList();
        Console.WriteLine("Notebook elements matched: " + nbs.Count);
        foreach (var nb in nbs) {
            string name = (string)nb.Attribute("name") ?? "";
            var groups = nb.Elements().Where(e => e.Name.LocalName == "SectionGroup").ToList();
            Console.WriteLine($"  NB '{name}': direct SectionGroup children = {groups.Count}");
            foreach (var g in groups)
                Console.WriteLine($"     group: {g.Attribute("name")}");
        }
    }
}
