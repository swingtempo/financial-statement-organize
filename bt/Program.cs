using System;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Interop.OneNote;

class P
{
    static void Main()
    {
        var appType = typeof(IApplication);
        Console.WriteLine("=== IApplication methods (with param names) ===");
        foreach (var m in appType.GetMethods().OrderBy(m => m.Name))
        {
            if (m.DeclaringType != appType) continue;
            var ps = m.GetParameters();
            string sig = m.Name + "(" + string.Join(", ", ps.Select(p => p.ParameterType.Name + " " + p.Name)) + ")";
            Console.WriteLine("  " + sig);
        }
        Console.WriteLine();
        Console.WriteLine("=== MergeFiles detail ===");
        foreach (var m in appType.GetMethods().Where(m => m.Name == "MergeFiles"))
        {
            Console.WriteLine("  " + m.ToString());
            foreach (var p in m.GetParameters())
                Console.WriteLine("    param: " + p.ParameterType + " " + p.Name + " (in=" + p.IsIn + ", out=" + p.IsOut + ")");
        }
    }
}
