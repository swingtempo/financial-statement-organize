using System;
using System.Text;

namespace OneNoteSync
{
    /// <summary>
    /// Builds the OneNote page-body XML for a statement page:
    ///   - summary text (one <one:OE> per line)
    ///   - the PDF attached as <one:InsertedFile pathSource=...> (OneNote ingests it)
    ///   - each rendered page image as <one:Image> with inline base64 <one:Data>
    ///
    /// This is the proven in-process write path: the block is inserted before
    /// </one:Page> and committed with UpdatePageContent.
    /// </summary>
    public static class PageXml
    {
        /// <summary>Builds the &lt;one:Outline&gt; body to insert before &lt;/one:Page&gt;.</summary>
        public static string BuildBody(string summary, string pdfPath, string pdfName,
                                      System.Collections.Generic.List<(string Name, byte[] Png)> rasters, int dpi)
        {
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.000Z");
            var sb = new StringBuilder();
            sb.Append("<one:Outline><one:OEChildren>");

            // 1. Summary text (one OE per non-empty line)
            if (!string.IsNullOrWhiteSpace(summary))
            {
                foreach (var line in summary.Split('\n'))
                {
                    var t = line.TrimEnd('\r');
                    if (t.Length == 0) continue;
                    sb.Append("<one:OE creationTime=\"").Append(now).Append("\" lastModifiedTime=\"")
                      .Append(now).Append("\" alignment=\"left\">")
                      .Append("<one:T>").Append(Esc(t)).Append("</one:T>")
                      .Append("</one:OE>");
                }
                // a spacer line
                sb.Append("<one:OE creationTime=\"").Append(now).Append("\" lastModifiedTime=\"")
                  .Append(now).Append("\" alignment=\"left\"><one:T> </one:T></one:OE>");
            }

            // 2. PDF attached as InsertedFile (pathSource; preferredName = bare filename)
            if (!string.IsNullOrEmpty(pdfPath))
            {
                var fn = (pdfName ?? "").Replace("\\", "").Replace("/", "").Trim();
                if (string.IsNullOrEmpty(fn)) fn = System.IO.Path.GetFileName(pdfPath);
                if (string.IsNullOrEmpty(fn)) fn = "document.pdf";
                sb.Append("<one:OE creationTime=\"").Append(now).Append("\" lastModifiedTime=\"")
                  .Append(now).Append("\" alignment=\"left\">")
                  .Append("<one:InsertedFile pathSource=\"").Append(Esc(pdfPath))
                  .Append("\" preferredName=\"").Append(Esc(fn)).Append("\"/>")
                  .Append("</one:OE>");
            }

            // 3. Page images as one:Image with inline base64 Data
            if (rasters != null)
            {
                foreach (var r in rasters)
                {
                    var (w, h) = PngSize(r.Png);
                    double pw = w * 72.0 / dpi;   // points
                    double ph = h * 72.0 / dpi;
                    string b64 = Convert.ToBase64String(r.Png);
                    sb.Append("<one:OE creationTime=\"").Append(now).Append("\" lastModifiedTime=\"")
                      .Append(now).Append("\" alignment=\"left\">")
                      .Append("<one:Image isPrintOut=\"true\" backgroundImage=\"true\" format=\"auto\">")
                      .Append("<one:Size width=\"").Append(pw.ToString("0.00")).Append("\" height=\"")
                      .Append(ph.ToString("0.00")).Append("\"/>")
                      .Append("<one:Data>").Append(b64).Append("</one:Data>")
                      .Append("</one:Image></one:OE>");
                }
            }

            sb.Append("</one:OEChildren></one:Outline>");
            return sb.ToString();
        }

        /// <summary>Reads width/height from a PNG's IHDR chunk.</summary>
        public static (int w, int h) PngSize(byte[] png)
        {
            // bytes 16-19 = width (big-endian), 20-23 = height
            int w = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
            int h = (png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23];
            return (w, h);
        }

        static string Esc(string s) => (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
    }
}
