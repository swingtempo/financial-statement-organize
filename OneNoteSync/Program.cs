using System.Windows.Forms;

namespace OneNoteSync;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        // OneNoteSync is Windows-only (OneNote desktop COM API).
        if (!System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
        {
            MessageBox.Show("OneNoteSync is Windows-only (OneNote desktop COM API).", "OneNoteSync",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm(args));
    }
}
