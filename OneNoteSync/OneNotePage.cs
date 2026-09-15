using System;
using System.Collections.Generic;
using System.Text;

namespace OneNoteSync
{
    /// <summary>
    /// Builds the OneNote page-body XML for a statement page:
    ///   - the page title in a &lt;one:Title&gt; element (precedes the body, per the OneNote schema)
    ///   - a summary table, one row per account: Date | Account | Balance | Notes
    ///   - the PDF attached as &lt;one:InsertedFile pathSource=...&gt; (OneNote ingests it)
    ///   - each rendered page image as &lt;one:Image&gt; with inline base64 &lt;one:Data&gt;
    ///
    /// This is the proven in-process write path: the block is inserted before
    /// &lt;/one:Page&gt; and committed with UpdatePageContent.
    /// </summary>
    public static class PageXml
    {
        /// <summary>One account (statement) row for the summary table.</summary>
        public sealed class AccountRow
        {
            public string Date { get; set; } = "";
            public string Name { get; set; } = "";
            public string Balance { get; set; } = "";
            public List<string> Notes { get; set; } = new();
        }

        // Column widths (points) for Date | Account | Balance | Notes.
        static readonly double[] ColWidths = { 90, 200, 110, 240 };
        static readonly string[] ColHeaders = { "Date", "Account", "Balance", "Notes" };

        /// <summary>Builds the &lt;one:Outline&gt; body to insert before &lt;/one:Page&gt;.</summary>
        public static string BuildBody(string title, List<AccountRow> rows,
                                      string pdfPath, string pdfName,
                                      List<(string Name, byte[] Png)> rasters, int dpi)
        {
            string now = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.000Z");
            var sb = new StringBuilder();

            // 1. The page title lives in <one:Title>, which the OneNote schema requires to
            //    precede all body elements (Outline/Image/InsertedFile/...).
            sb.Append(BuildTitle(now, title));

            // 2. The body: a single Outline holding the table, spacer, PDF, and images.
            sb.Append("<one:Outline><one:OEChildren>");

            // Summary table (one row per account).
            if (rows != null && rows.Count > 0)
                sb.Append(BuildTable(now, rows));

            // A spacer line between the table and the attached file/images.
            sb.Append(OE(now, " "));

            // 3. PDF attached as InsertedFile (pathSource; preferredName = bare filename)
            if (!string.IsNullOrEmpty(pdfPath))
            {
                var fn = (pdfName ?? "").Replace("\\", "").Replace("/", "").Trim();
                if (string.IsNullOrEmpty(fn)) fn = System.IO.Path.GetFileName(pdfPath);
                if (string.IsNullOrEmpty(fn)) fn = "document.pdf";
                sb.Append(OEOpen(now));
                sb.Append("<one:InsertedFile pathSource=\"").Append(Esc(pdfPath))
                  .Append("\" preferredName=\"").Append(Esc(fn)).Append("\"/>");
                sb.Append("</one:OE>");
            }

            // 4. Page images as one:Image with inline base64 Data
            if (rasters != null)
            {
                foreach (var r in rasters)
                {
                    var (w, h) = PngSize(r.Png);
                    double pw = w * 72.0 / dpi;   // points
                    double ph = h * 72.0 / dpi;
                    string b64 = Convert.ToBase64String(r.Png);
                    sb.Append(OEOpen(now));
                    sb.Append("<one:Image isPrintOut=\"true\" backgroundImage=\"true\" format=\"auto\">")
                      .Append("<one:Size width=\"").Append(pw.ToString("0.00")).Append("\" height=\"")
                      .Append(ph.ToString("0.00")).Append("\"/>")
                      .Append("<one:Data>").Append(b64).Append("</one:Data>")
                      .Append("</one:Image>");
                    sb.Append("</one:OE>");
                }
            }

            sb.Append("</one:OEChildren></one:Outline>");
            return sb.ToString();
        }

        /// <summary>
        /// Builds the &lt;one:Table&gt; (wrapped in an outer &lt;one:OE&gt;). One header row,
        /// then one data row per account. Each cell wraps its &lt;one:OE&gt; in
        /// &lt;one:OEChildren&gt;; the Notes cell carries one OE per verbatim note line.
        /// </summary>
        static string BuildTable(string now, List<AccountRow> rows)
        {
            var sb = new StringBuilder();
            sb.Append(OEOpen(now));
            sb.Append("<one:Table bordersVisible=\"true\" hasHeaderRow=\"true\">");

            // Columns (index + width are required).
            sb.Append("<one:Columns>");
            for (int i = 0; i < ColWidths.Length; i++)
                sb.Append("<one:Column index=\"").Append(i).Append("\" width=\"")
                  .Append(ColWidths[i].ToString("0.00")).Append("\"/>");
            sb.Append("</one:Columns>");

            // Header row.
            sb.Append("<one:Row>");
            for (int i = 0; i < ColHeaders.Length; i++)
                sb.Append(Cell(now, ColHeaders[i]));
            sb.Append("</one:Row>");

            // Data rows (one per account).
            foreach (var r in rows)
            {
                sb.Append("<one:Row>");
                sb.Append(Cell(now, r.Date));
                sb.Append(Cell(now, r.Name));
                sb.Append(Cell(now, r.Balance));
                // Notes: one OE per verbatim note line (empty OE when there are none).
                sb.Append("<one:Cell><one:OEChildren>");
                if (r.Notes != null && r.Notes.Count > 0)
                    foreach (var n in r.Notes)
                        sb.Append(OE(now, n));
                else
                    sb.Append(OE(now, ""));
                sb.Append("</one:OEChildren></one:Cell>");
                sb.Append("</one:Row>");
            }

            sb.Append("</one:Table>");
            sb.Append("</one:OE>");
            return sb.ToString();
        }

        /// <summary>
        /// Builds the &lt;one:Title&gt; element (the page title). Per the OneNote schema it must
        /// appear before any body element and appears at most once.
        /// </summary>
        static string BuildTitle(string now, string title)
        {
            var t = title ?? "";
            return "<one:Title lang=\"en-US\">"
                 + "<one:OE creationTime=\"" + now + "\" lastModifiedTime=\"" + now + "\" alignment=\"left\">"
                 + "<one:T>" + Esc(t) + "</one:T>"
                 + "</one:OE>"
                 + "</one:Title>";
        }

        static string Cell(string now, string text)
            => "<one:Cell><one:OEChildren>" + OE(now, text) + "</one:OEChildren></one:Cell>";

        static string OEOpen(string now)
            => "<one:OE creationTime=\"" + now + "\" lastModifiedTime=\"" + now + "\" alignment=\"left\">";

        static string OE(string now, string text)
            => OEOpen(now) + "<one:T>" + Esc(text) + "</one:T></one:OE>";

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
