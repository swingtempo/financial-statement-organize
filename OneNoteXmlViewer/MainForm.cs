using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using Patagames.Pdf;
using Patagames.Pdf.Net;
using Patagames.Pdf.Enums;

namespace OneNoteXmlViewer
{
    class PageItem
    {
        public string Display; public string Id;
        public PageItem(string d, string i) { Display = d; Id = i; }
        public override string ToString() => Display;
    }

    public class MainForm : Form
    {
        Microsoft.Office.Interop.OneNote.Application _app;
        TextBox txtSearch, txtPageId, txtXml, txtBinInfo, txtLog;
        TextBox txtPdfPath, txtPreferredName, txtInsertResult;
        CheckBox chkImages;
        ComboBox cboPages, cboPageInfo, cboWriteMethod;
        ListBox lstCallbacks;
        List<PageItem> _allPages = new List<PageItem>();
        string _currentPageId = "";
        string _currentXml = "";
        byte[] _currentBin = null;
        const string NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        public MainForm()
        {
            Text = "OneNote XML Viewer  (selectable PageInfo) + PDF Insert (pathSource) Test";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1160, 850);
            Font = new Font("Segoe UI", 9f);
            System.Windows.Forms.Application.ThreadException += (s, e) => LogException(e.Exception, "ThreadException");
            System.AppDomain.CurrentDomain.UnhandledException += (s, e) => LogException(e.ExceptionObject as Exception, "UnhandledException");
            InitUI();
            BeginInit();
        }

        void InitUI()
        {
            // Row 1
            txtSearch = TB(8, 8, 280); Controls.Add(txtSearch);
            Controls.Add(Btn("Find pages", 292, 8, 90, (s, e) => FindPages()));
            txtPageId = TB(388, 8, 300); Controls.Add(txtPageId);
            Controls.Add(Btn("Load by ID", 694, 8, 90, (s, e) => LoadPage()));
            Controls.Add(Lbl("PageInfo:", 790, 12));
            cboPageInfo = new ComboBox { Location = new Point(866, 8), Width = 240, DropDownStyle = ComboBoxStyle.DropDownList };
            Controls.Add(cboPageInfo);
            foreach (PageInfo pi in Enum.GetValues(typeof(PageInfo)))
                cboPageInfo.Items.Add(pi.ToString() + "  (" + (int)pi + ")");
            cboPageInfo.SelectedIndex = 0; // piBasic (default)
            cboPageInfo.SelectedIndexChanged += (s, e) => { string v = cboPageInfo.SelectedItem != null ? cboPageInfo.SelectedItem.ToString() : "none"; Log("PageInfo set to: " + v); };

            // Row 2
            Controls.Add(Lbl("Page:", 8, 42));
            cboPages = new ComboBox { Location = new Point(56, 40), Width = 1096, DropDownStyle = ComboBoxStyle.DropDownList };
            Controls.Add(cboPages);
            cboPages.SelectedIndexChanged += (s, e) => { if (cboPages.SelectedItem is PageItem pi) txtPageId.Text = pi.Id; };

            // XML
            Controls.Add(Lbl("Page XML:", 8, 70));
            txtXml = TB(8, 88, 1144, 250, mono: true); Controls.Add(txtXml);

            // Button row
            int by = 344;
            Controls.Add(Btn("Refresh pages", 8, by, 120, (s, e) => RefreshPages()));
            Controls.Add(Btn("Save XML", 134, by, 110, (s, e) => SaveXml()));
            Controls.Add(Btn("Save binary", 250, by, 110, (s, e) => SaveBin()));
            Controls.Add(Btn("Save all binaries", 366, by, 150, (s, e) => SaveAllBin()));
            Controls.Add(Btn("Write back (current page)", 522, by, 210, (s, e) => WriteBack()));
            Controls.Add(Btn("Show skeleton", 738, by, 150, (s, e) => ShowSkeleton()));
            Controls.Add(Lbl("Write via:", 896, by + 4));
            cboWriteMethod = new ComboBox { Location = new Point(968, by), Width = 178, DropDownStyle = ComboBoxStyle.DropDownList };
            cboWriteMethod.Items.Add("PIA direct (1-arg)");
            cboWriteMethod.Items.Add("COM dynamic (1-arg)");
            cboWriteMethod.Items.Add("Sidecar (PowerShell)");
            cboWriteMethod.SelectedIndex = 0;
            Controls.Add(cboWriteMethod);

            // Binary
            Controls.Add(Lbl("Binary parts (CallbackIDs):", 8, 378));
            lstCallbacks = new ListBox { Location = new Point(8, 396), Size = new Size(600, 90) }; Controls.Add(lstCallbacks);
            lstCallbacks.SelectedIndexChanged += (s, e) => { if (lstCallbacks.SelectedItem != null) FetchBinary(); };
            Controls.Add(Btn("Fetch binary", 614, 396, 130, (s, e) => FetchBinary()));
            Controls.Add(Lbl("Binary info:", 760, 378));
            txtBinInfo = TB(760, 396, 392, 90, mono: true); Controls.Add(txtBinInfo);

            // PDF INSERT TEST
            Controls.Add(Lbl("PDF insert test — inject <one:InsertedFile pathSource=...> and let OneNote ingest it:", 8, 494));
            Controls.Add(Lbl("pathSource (file):", 8, 514));
            txtPdfPath = TB(150, 512, 640); Controls.Add(txtPdfPath);
            Controls.Add(Btn("Browse...", 796, 510, 80, (s, e) => BrowsePdf()));
            Controls.Add(Lbl("preferredName:", 884, 514));
            txtPreferredName = TB(984, 512, 168); Controls.Add(txtPreferredName);
            Controls.Add(Btn("Show XML only", 8, 548, 140, (s, e) => ShowInsertXml()));
            Controls.Add(Btn("Insert into page + verify", 154, 548, 200, (s, e) => DoInsert()));
            chkImages = new CheckBox { Text = "Also insert each page as an image", Location = new Point(370, 548), AutoSize = true, Checked = true };
            Controls.Add(chkImages);
            Controls.Add(Lbl("Result:", 8, 582));
            txtInsertResult = TB(8, 600, 1144, 60, mono: true); Controls.Add(txtInsertResult);

            // Log
            Controls.Add(Lbl("Log:", 8, 668));
            txtLog = TB(8, 686, 1144, 140, mono: true); Controls.Add(txtLog);
        }

        void BeginInit()
        {
            try
            {
                _app = new ApplicationClass();
                Log("PIA ApplicationClass created OK");
            }
            catch (Exception e) { Log("PIA create FAIL: " + Err(e)); return; }
            EnsureOneNote();
            RefreshPages();
        }

        // ---------- helpers ----------
        string Err(Exception e)
        {
            if (e is System.Runtime.InteropServices.COMException ce)
                return "0x" + ce.HResult.ToString("X8") + " (" + ce.Message + ")";
            return e.Message;
        }

        void LogException(Exception ex, string source = "EXCEPTION")
        {
            if (ex == null) { Log(source + ": (null)"); return; }
            Log(source + ": " + ex.ToString());
            try { File.AppendAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "onenoteviewer_errors.log"), "[" + DateTime.Now.ToString("s") + "] " + source + " " + ex.ToString() + "\r\n\r\n"); } catch { }
        }

        void Log(string m)
        {
            if (txtLog == null) return;
            if (txtLog.InvokeRequired) { txtLog.BeginInvoke((Action)(() => AppendLog(m))); return; }
            AppendLog(m);
        }
        void AppendLog(string m) { txtLog.AppendText("[ " + DateTime.Now.ToString("HH:mm:ss") + " ] " + m + Environment.NewLine); }

        void SetResult(string m)
        {
            if (txtInsertResult == null) return;
            if (txtInsertResult.InvokeRequired) { txtInsertResult.BeginInvoke((Action)(() => AppendResult(m))); return; }
            AppendResult(m);
        }
        void AppendResult(string m) { txtInsertResult.AppendText(m + Environment.NewLine); }

        void EnsureOneNote()
        {
            var p = Process.GetProcessesByName("ONENOTE");
            if (p.Length > 0) { Log("OneNote already running (pid " + p[0].Id + ")"); p[0].Dispose(); return; }
            string exe = @"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE";
            if (File.Exists(exe)) { try { Process.Start(exe); Log("Launched OneNote"); Thread.Sleep(4000); } catch (Exception e) { Log("Launch fail: " + e.Message); } }
            else Log("OneNote not running; exe not found at default path");
        }

        string XmlEsc(string s) => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

        PageInfo GetPageInfo()
        {
            if (cboPageInfo != null && cboPageInfo.SelectedItem != null)
            {
                string s = cboPageInfo.SelectedItem.ToString(); // e.g. "piAll  (7)"
                int p = s.IndexOf(" (");
                if (p > 0) s = s.Substring(0, p);
                PageInfo val;
                if (Enum.TryParse(s, out val)) return val;
            }
            return PageInfo.piAll;
        }

        // ---------- pages ----------
        void RefreshPages()
        {
            try
            {
                string hier = "";
                _app.GetHierarchy("", HierarchyScope.hsPages, out hier);
                var doc = XDocument.Parse(hier);
                var ns = XNamespace.Get(NS);
                _allPages.Clear();
                int n = 0;
                foreach (var page in doc.Descendants(ns + "Page"))
                {
                    string name = (string)page.Attribute("name") ?? "";
                    string id = (string)page.Attribute("ID") ?? "";
                    var sec = page.AncestorsAndSelf(ns + "Section").FirstOrDefault();
                    var nb = page.AncestorsAndSelf(ns + "Notebook").FirstOrDefault();
                    string secN = sec != null ? (string)sec.Attribute("name") : "";
                    string nbN = nb != null ? (string)nb.Attribute("name") : "";
                    _allPages.Add(new PageItem(name + "   —   " + secN + " / " + nbN, id));
                    n++;
                }
                FillCombo(_allPages);
                Log("Refresh: " + n + " pages loaded");
            }
            catch (Exception e) { Log("GetHierarchy FAIL: " + Err(e)); }
        }

        void FindPages()
        {
            string q = txtSearch.Text.Trim().ToLower();
            var match = _allPages.Count == 0 ? _allPages : _allPages.Where(p => p.Display.ToLower().Contains(q)).ToList();
            FillCombo(match);
            Log("Find '" + txtSearch.Text + "': " + match.Count + " matches");
        }

        void FillCombo(List<PageItem> list)
        {
            cboPages.Items.Clear();
            foreach (var p in list) cboPages.Items.Add(p);
            if (list.Count > 0) cboPages.SelectedIndex = 0;
        }

        // ---------- load page ----------
        void LoadPage()
        {
            string pid = txtPageId.Text.Trim();
            if (string.IsNullOrEmpty(pid) && cboPages.SelectedItem is PageItem pi) pid = pi.Id;
            if (string.IsNullOrEmpty(pid)) { Log("No page ID. Pick a page or enter an ID."); return; }
            _currentPageId = pid;
            _currentXml = "";
            try
            {
                string xml = "";
                _app.GetPageContent(pid, out xml, GetPageInfo());
                _currentXml = xml;
                txtXml.Text = xml;
                Log("Loaded " + pid + "  xml len=" + xml.Length + "  (PageInfo=" + GetPageInfo() + ")");
                ExtractCallbacks(xml);
            }
            catch (Exception e) { Log("GetPageContent FAIL: " + Err(e)); }
        }

        void ExtractCallbacks(string xml)
        {
            lstCallbacks.Items.Clear();
            var m = Regex.Matches(xml, "callbackID=\"([^\"]+)\"");
            foreach (Match mm in m)
            {
                string cid = mm.Groups[1].Value;
                if (!lstCallbacks.Items.Contains(cid)) lstCallbacks.Items.Add(cid);
            }
            Log("Found " + lstCallbacks.Items.Count + " binary CallbackIDs in XML");
        }

        // ---------- binary ----------
        void FetchBinary()
        {
            if (string.IsNullOrEmpty(_currentPageId)) { Log("Load a page first."); return; }
            if (lstCallbacks.SelectedItem == null) { Log("Select a CallbackID."); return; }
            string cid = lstCallbacks.SelectedItem.ToString();
            string b64 = "";
            try
            {
                _app.GetBinaryPageContent(_currentPageId, cid, out b64);
                _currentBin = Convert.FromBase64String(b64);
                txtBinInfo.Text = DescribeBinary(_currentBin);
                Log("Fetched " + cid + "  len=" + _currentBin.Length + "  " + TypeOf(_currentBin));
            }
            catch (Exception e) { Log("GetBinaryPageContent FAIL: " + Err(e)); }
        }

        string DescribeBinary(byte[] b)
        {
            if (b == null) return "";
            var sb = new StringBuilder();
            sb.AppendLine("Size: " + b.Length + " bytes");
            sb.AppendLine("Type: " + TypeOf(b));
            sb.Append("Magic: " + Hex(b, 16));
            return sb.ToString();
        }

        string TypeOf(byte[] b)
        {
            if (b.Length >= 4 && b[0] == 0x25 && b[1] == 0x50 && b[2] == 0x44 && b[3] == 0x46) return "%PDF (PDF file)";
            if (b.Length >= 3 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF) return "JPEG";
            if (b.Length >= 4 && b[0] == 0x89 && b[1] == 0x50 && b[2] == 0x4E && b[3] == 0x47) return "PNG";
            if (b.Length >= 2 && b[0] == 0x47 && b[1] == 0x49) return "GIF";
            if (b.Length >= 4 && b[0] == 0x42 && b[1] == 0x4D) return "BMP";
            return "unknown";
        }

        string Hex(byte[] b, int n)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < Math.Min(n, b.Length); i++) sb.Append(b[i].ToString("X2")).Append(' ');
            return sb.ToString();
        }

        // ---------- save ----------
        void ShowSkeleton()
        {
            if (string.IsNullOrEmpty(_currentXml)) { Log("No XML loaded."); return; }
            string skeleton = Regex.Replace(_currentXml, "[A-Za-z0-9+/=]{400,}", m => "[...base64 " + m.Length + " chars...]");
            txtXml.Text = skeleton;
            Log("Show skeleton (base64 truncated) -> " + skeleton.Length + " chars");
        }

        void SaveXml()
        {
            if (string.IsNullOrEmpty(_currentXml)) { Log("No XML loaded."); return; }
            using (var sfd = new SaveFileDialog { Filter = "XML|*.xml", FileName = "page.xml" })
            { if (sfd.ShowDialog() != DialogResult.OK) return; File.WriteAllText(sfd.FileName, _currentXml); Log("Saved XML -> " + sfd.FileName); }
        }

        void SaveBin()
        {
            if (_currentBin == null) { Log("No binary loaded. Fetch one first."); return; }
            using (var sfd = new SaveFileDialog { Filter = "All|*.*", FileName = "part.bin" })
            { if (sfd.ShowDialog() != DialogResult.OK) return; File.WriteAllBytes(sfd.FileName, _currentBin); Log("Saved binary -> " + sfd.FileName); }
        }

        void SaveAllBin()
        {
            if (lstCallbacks.Items.Count == 0) { Log("No CallbackIDs."); return; }
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() != DialogResult.OK) return;
                string dir = fbd.SelectedPath;
                int n = 0;
                foreach (System.Object cidObj in lstCallbacks.Items)
                {
                    string cid = cidObj.ToString();
                    string b64 = "";
                    try
                    {
                        _app.GetBinaryPageContent(_currentPageId, cid, out b64);
                        byte[] b = Convert.FromBase64String(b64);
                        string t = TypeOf(b);
                        string ext = t.Contains("PDF") ? ".pdf" : (t.Contains("JPEG") ? ".jpg" : (t.Contains("PNG") ? ".png" : ".bin"));
                        string safe = cid.Replace("{", "").Replace("}", "").Replace("-", "");
                        File.WriteAllBytes(Path.Combine(dir, safe + ext), b);
                        n++;
                    }
                    catch (Exception e) { Log("  fail " + cid + ": " + Err(e)); }
                }
                Log("Saved " + n + " binaries -> " + dir);
            }
        }

        // ---------- sidecar write (1-arg, current page) ----------
        bool SidecarUpdate(string xml)
        {
            try
            {
                string tmp = Path.Combine(Path.GetTempPath(), "oneview_" + Guid.NewGuid().ToString() + ".xml");
                File.WriteAllText(tmp, xml);
                string ps = "$a = New-Object -ComObject OneNote.Application\r\n$x = Get-Content -Raw -LiteralPath '" + tmp + "'\r\n$a.UpdatePageContent($x)\r\n";
                string psf = Path.Combine(Path.GetTempPath(), "oneview_" + Guid.NewGuid().ToString() + ".ps1");
                File.WriteAllText(psf, ps);
                var p = Process.Start("powershell.exe", "-NoProfile -ExecutionPolicy Bypass -File \"" + psf + "\"");
                p.WaitForExit();
                Log("Sidecar UpdatePageContent exit=" + p.ExitCode);
                return p.ExitCode == 0;
            }
            catch (Exception e) { Log("Sidecar FAIL: " + e.Message); return false; }
        }

        // ---------- direct write (PIA / raw COM) ----------
        string UpdateVia(string xml)
        {
            string mode = cboWriteMethod != null && cboWriteMethod.SelectedItem != null ? cboWriteMethod.SelectedItem.ToString() : "PIA direct";
            try
            {
                if (mode.StartsWith("PIA"))
                {
                    _app.UpdatePageContent(xml);   // PIA 1-arg, in-process
                    return "PIA direct OK";
                }
                if (mode.StartsWith("COM"))
                {
                    object raw = Activator.CreateInstance(Type.GetTypeFromProgID("OneNote.Application"));
                    dynamic dyn = raw;
                    dyn.UpdatePageContent(xml);   // raw COM 1-arg
                    return "COM dynamic OK";
                }
                return SidecarUpdate(xml) ? "Sidecar OK" : "Sidecar FAIL";
            }
            catch (Exception e)
            {
                LogException(e, "UpdateVia FAIL");
                return "FAIL " + Err(e);
            }
        }

        void WriteBack()
        {
            if (string.IsNullOrEmpty(_currentXml)) { Log("No XML loaded."); return; }
            if (MessageBox.Show("Write the displayed XML back to the page open in OneNote?\n(Confirmed 1-arg sidecar path. Target page made current first.)",
                "Confirm write-back", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            try
            {
                try { _app.NavigateTo(_currentPageId); Log("NavigateTo (make current)"); }
                catch (Exception e) { Log("NavigateTo warn: " + Err(e)); }
                Thread.Sleep(800);
                _currentXml = txtXml.Text;
                string res = UpdateVia(_currentXml);
                Log("Write-back: " + res);
            }
            catch (Exception e) { LogException(e, "Write-back FAIL"); }
        }

        void BrowsePdf()
        {
            using (var ofd = new OpenFileDialog { Filter = "PDF|*.pdf|All|*.*" })
            { if (ofd.ShowDialog() == DialogResult.OK) { txtPdfPath.Text = ofd.FileName; if (string.IsNullOrEmpty(txtPreferredName.Text)) txtPreferredName.Text = Path.GetFileName(ofd.FileName); } }
        }

        // ---------- PDF INSERT TEST ----------
        string BuildInsertedFileBlock(string pdf, string name, List<(string Path, double W, double H)> images)
        {
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.000Z");

            // preferredName MUST NOT contain backslashes (or slashes) — bare filename only
            name = (name ?? "").Replace("\\", "").Replace("/", "").Trim();
            if (string.IsNullOrEmpty(name)) name = "document.pdf";

            string ifAttrs = "<one:InsertedFile pathSource=\"" + XmlEsc(pdf) + "\" preferredName=\"" + XmlEsc(name) + "\"/>";

            string oes = "<one:OE creationTime=\"" + now + "\" lastModifiedTime=\"" + now + "\" alignment=\"left\">" +
                       ifAttrs +
                       "</one:OE>";

            // Page images as sibling OEs under the same outline.
            // <one:Image> supplies its binary via <one:Data> (inline base64),
            // with <one:Size> first (PageObject base), then the Data element.
            if (images != null)
            {
                foreach (var img in images)
                {
                    string b64 = Convert.ToBase64String(File.ReadAllBytes(img.Path));
                    string w = img.W.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
                    string h = img.H.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
                    oes += "<one:OE creationTime=\"" + now + "\" lastModifiedTime=\"" + now + "\" alignment=\"left\">" +
                         "<one:Image isPrintOut=\"true\" backgroundImage=\"true\" format=\"auto\">" +
                         "<one:Size width=\"" + w + "\" height=\"" + h + "\"/>" +
                         "<one:Data>" + b64 + "</one:Data>" +
                         "</one:Image></one:OE>";
                }
                Log("Emitting " + images.Count + " page-image one:Image OE(s) with inline base64 <one:Data>");
            }

            return "<one:Outline><one:OEChildren>" + oes + "</one:OEChildren></one:Outline>";
        }

        // ---------- PDF PAGE-IMAGE RASTERIZATION (PDFium) ----------
        List<(string Name, byte[] Png)> Rasterize(string pdfPath, int dpi, int maxPages)
        {
            var pages = new List<(string, byte[])>();
            using (var doc = PdfDocument.Load(pdfPath, null, null))
            {
                int count = Math.Min(doc.Pages.Count, maxPages);
                for (int i = 0; i < count; i++)
                {
                    var pg = doc.Pages[i];
                    var sz = doc.GetPageSizeByIndex(i);
                    double scale = dpi / 72.0;
                    int w = Math.Max(1, (int)(sz.Width * scale));
                    int h = Math.Max(1, (int)(sz.Height * scale));
                    using (var bmp = new PdfBitmap(w, h, true))
                    {
                        bmp.FillRect(0, 0, w, h, FS_COLOR.White);
                        pg.Render(bmp, 0, 0, w, h, PageRotate.Normal, RenderFlags.FPDF_NONE);
                        int stride = bmp.Stride;
                        var raw = new byte[h * stride];
                        System.Runtime.InteropServices.Marshal.Copy(bmp.Buffer, raw, 0, raw.Length);
                        var rows = new MemoryStream();
                        for (int y = 0; y < h; y++) { rows.WriteByte(0); rows.Write(raw, y * stride, stride); }
                        pages.Add(("page" + (i + 1), PngWriter.Encode(rows.ToArray(), w, h)));
                    }
                }
            }
            return pages;
        }

        // Render all PDF pages to PNG files in a temp dir; return (file, width_pt, height_pt).
        List<(string Path, double W, double H)> RenderPageImages(string pdf)
        {
            const int dpi = 150;
            var pages = Rasterize(pdf, dpi, 9999);
            var dir = Path.Combine(Path.GetTempPath(), "oneview_img_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            var result = new List<(string Path, double W, double H)>();
            double scale = dpi / 72.0;
            foreach (var p in pages)
            {
                string file = Path.Combine(dir, p.Name + ".png");
                File.WriteAllBytes(file, p.Png);
                int pw, ph;
                using (var bmp = new System.Drawing.Bitmap(new MemoryStream(p.Png))) { pw = bmp.Width; ph = bmp.Height; }
                double w = pw / scale, h = ph / scale;
                result.Add((file, w, h));
                Log(string.Format("Rendered {0}: {1}x{2}px -> one:Size {3:F0}x{4:F0}pt  [{5}]", p.Name, pw, ph, w, h, file));
            }
            return result;
        }

        // ---------- PNG ENCODER (8-bit RGBA, filter 0) ----------
        static class PngWriter
        {
            public static byte[] Encode(byte[] rows, int width, int height)
            {
                using var idatMs = new MemoryStream();
                using (var deflate = new System.IO.Compression.DeflateStream(idatMs, System.IO.Compression.CompressionLevel.Optimal, leaveOpen: true))
                    deflate.Write(rows, 0, rows.Length);
                var rawDeflate = idatMs.ToArray();
                long adler = Adler32(rows);
                var zlibPayload = new byte[rawDeflate.Length + 6];
                zlibPayload[0] = 0x78; zlibPayload[1] = 0x9C;
                Buffer.BlockCopy(rawDeflate, 0, zlibPayload, 2, rawDeflate.Length);
                zlibPayload[zlibPayload.Length - 4] = (byte)(adler >> 24); zlibPayload[zlibPayload.Length - 3] = (byte)(adler >> 16);
                zlibPayload[zlibPayload.Length - 2] = (byte)(adler >> 8);  zlibPayload[zlibPayload.Length - 1] = (byte)adler;

                using var ms = new MemoryStream();
                void W(byte[] b) => ms.Write(b, 0, b.Length);
                void WBE32(int v) { W(new byte[] { (byte)(v >> 24), (byte)(v >> 16), (byte)(v >> 8), (byte)v }); }
                void Chunk(string type, byte[] data)
                {
                    WBE32(data.Length);
                    var tb = System.Text.Encoding.ASCII.GetBytes(type);
                    W(tb); W(data);
                    WBE32((int)Crc32(tb.Concat(data).ToArray()));
                }
                W(new byte[] { 0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A });
                var ihdr = new byte[13];
                for (int k = 0; k < 4; k++) { ihdr[k] = (byte)(width >> (24 - 8 * k)); ihdr[4 + k] = (byte)(height >> (24 - 8 * k)); }
                ihdr[8] = 8; ihdr[9] = 6;
                Chunk("IHDR", ihdr);
                Chunk("IDAT", zlibPayload);
                Chunk("IEND", Array.Empty<byte>());
                return ms.ToArray();
            }
            static long Adler32(byte[] data) { uint a = 1, b = 0; foreach (var x in data) { a = (a + x) % 65521; b = (b + a) % 65521; } return (b << 16) | a; }
            static uint Crc32(byte[] data) { uint crc = 0xFFFFFFFF; foreach (var b in data) { crc ^= b; for (int i = 0; i < 8; i++) crc = (crc >> 1) ^ (0xEDB88320u & (uint)(-(crc & 1))); } return crc ^ 0xFFFFFFFF; }
        }

        void ShowInsertXml()
        {
            string pdf = txtPdfPath.Text.Trim();
            string name = string.IsNullOrEmpty(txtPreferredName.Text.Trim()) ? Path.GetFileName(pdf) : txtPreferredName.Text.Trim();
            if (string.IsNullOrEmpty(pdf)) { SetResult("Enter a pathSource file path first."); return; }
            SetResult("--- Injected block (not committed) ---");
            SetResult(BuildInsertedFileBlock(pdf, name, null));
            Log("ShowInsertXml (no commit)");
        }

        void DoInsert()
        {
            string pdf = txtPdfPath.Text.Trim();
            string name = string.IsNullOrEmpty(txtPreferredName.Text.Trim()) ? Path.GetFileName(pdf) : txtPreferredName.Text.Trim();
            if (string.IsNullOrEmpty(pdf)) { SetResult("Enter a pathSource file path."); return; }
            if (!File.Exists(pdf)) { SetResult("File not found: " + pdf); return; }
            if (string.IsNullOrEmpty(_currentPageId)) { SetResult("Load a target page first."); return; }

            int beforeIf = -1, afterIf = -1;
            string beforeCache = "", afterCache = "";
            try
            {
                // 1. Get current XML (load if needed)
                string xml = _currentXml;
                if (string.IsNullOrEmpty(xml))
                {
                    try { _app.GetPageContent(_currentPageId, out xml, GetPageInfo()); _currentXml = xml; }
                    catch (Exception e) { SetResult("GetPageContent fail: " + Err(e)); return; }
                }
                beforeIf = Regex.Matches(xml, "<one:InsertedFile").Count;
                beforeCache = Regex.Match(xml, "pathCache=\"([^\"]*)\"").Groups[1].Value;

                // 2. (Optional) render each PDF page to a PNG file
                List<(string Path, double W, double H)> images = null;
                if (chkImages != null && chkImages.Checked)
                {
                    try { images = RenderPageImages(pdf); }
                    catch (Exception e) { LogException(e, "RenderPageImages"); SetResult("Render fail: " + e.Message); return; }
                }

                // 3. Insert the InsertedFile + image block before </one:Page>
                string block = BuildInsertedFileBlock(pdf, name, images);
                int idx = xml.LastIndexOf("</one:Page>");
                if (idx < 0) { SetResult("Cannot find </one:Page> to insert into."); return; }
                string newXml = xml.Insert(idx, block);
                _currentXml = newXml;
                txtXml.Text = newXml;
                Log("Built new XML len=" + newXml.Length + " (inserted InsertedFile for " + name + ")");

                // 3. Make current + commit via sidecar
                try { _app.NavigateTo(_currentPageId); Log("NavigateTo (make current)"); } catch (Exception e) { Log("NavigateTo warn: " + Err(e)); }
                Thread.Sleep(1000);
                string wres = UpdateVia(newXml);
                bool ok = wres.EndsWith("OK");

                // 4. Wait for OneNote to ingest, then reload
                Log("Waiting 4s for OneNote to ingest the file...");
                Thread.Sleep(4000);
                string after = "";
                _app.GetPageContent(_currentPageId, out after, GetPageInfo());
                _currentXml = after;
                txtXml.Text = after;
                afterIf = Regex.Matches(after, "<one:InsertedFile").Count;
                var cm = Regex.Match(after, "pathCache=\"([^\"]*)\"");
                afterCache = cm.Success ? cm.Groups[1].Value : "";
                ExtractCallbacks(after);

                // 5. Report
                var sb = new StringBuilder();
                sb.AppendLine("write: " + wres);
                sb.AppendLine("InsertedFile count: " + beforeIf + " -> " + afterIf);
                sb.AppendLine("pathCache before: " + (string.IsNullOrEmpty(beforeCache) ? "(none)" : beforeCache));
                sb.AppendLine("pathCache after:  " + (string.IsNullOrEmpty(afterCache) ? "(none)" : afterCache));
                bool ingested = (afterIf > beforeIf) && !string.IsNullOrEmpty(afterCache);
                sb.AppendLine(ingested ? "RESULT: INGESTED (pathCache present) — OneNote cached the file!"
                                      : "RESULT: not confirmed ingested (check above)");
                SetResult(sb.ToString());
            }
            catch (Exception e) { LogException(e, "DoInsert FAIL"); SetResult("DoInsert FAIL: " + e.Message); }
        }

        // ---------- UI builders ----------
        TextBox TB(int x, int y, int w, int h = 26, bool mono = false) =>
            new TextBox { Location = new Point(x, y), Width = w, Height = h, Multiline = h > 40, ScrollBars = h > 40 ? ScrollBars.Both : ScrollBars.None, ReadOnly = false, Font = mono ? new Font("Consolas", 9f) : null, WordWrap = true };

        Button Btn(string text, int x, int y, int w, EventHandler handler)
        { var b = new Button { Text = text, Location = new Point(x, y), Size = new Size(w, 28) }; b.Click += handler; return b; }

        Label Lbl(string text, int x, int y) => new Label { Text = text, Location = new Point(x, y), AutoSize = true };
    }
}
