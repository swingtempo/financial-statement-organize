using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using StatementOrganizer;

var envPath = Path.GetFullPath(".env");
var projectEnvPath = Path.GetFullPath("../.env");
var env = EnvFile.Load(envPath, projectEnvPath);

// Resolve relative .env paths (INPUT_DIR/OUTPUT_DIR) against the .env file's folder.
string? envDir = null;
foreach (var p in new[] { envPath, projectEnvPath })
    if (File.Exists(p)) { envDir = Path.GetDirectoryName(p); break; }
static string ResolveDir(string? dir, string? baseDir) =>
    Path.GetFullPath(string.IsNullOrEmpty(dir)
        ? baseDir ?? "."
        : (Path.IsPathRooted(dir) ? dir : Path.Combine(baseDir ?? ".", dir)));

var apiKey = EnvFile.Get(env, "LLM_API_KEY")
    ?? throw new InvalidOperationException("LLM_API_KEY is not set (put it in .env).");
var baseUrl = (EnvFile.Get(env, "LLM_BASE_URL", "https://api.openai.com/v1") ?? "").TrimEnd('/');
var model = EnvFile.Get(env, "LLM_MODEL", "gpt-4o-mini")!;

var inputDir = args.Length > 0
    ? Path.GetFullPath(args[0])
    : ResolveDir(EnvFile.Get(env, "INPUT_DIR", "./input"), envDir);
var outputDir = args.Length > 1
    ? Path.GetFullPath(args[1])
    : ResolveDir(EnvFile.Get(env, "OUTPUT_DIR", "./organized"), envDir);

if (!Directory.Exists(inputDir))
    throw new DirectoryNotFoundException($"Input folder not found: {inputDir}");
Directory.CreateDirectory(outputDir);

var jsonOptions = new JsonSerializerOptions
{
    PropertyNameCaseInsensitive = true,
    WriteIndented = true,
};

var haveVision = PdfTools.HasRasterizer;
var haveText = PdfTools.HasTextExtractor;
var maxPages = int.Parse(EnvFile.Get(env, "MAX_PAGES", "20")!);
var dpi = int.Parse(EnvFile.Get(env, "PDF_DPI", "150")!);
Console.WriteLine($"Backends: files (upload) -> vision {(haveVision ? "ok" : "MISSING pdftoppm")} -> text {(haveText ? "ok" : "MISSING pdftotext")}");


using var http = new HttpClient { BaseAddress = new Uri(baseUrl + "/") };
http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

Console.WriteLine($"Model:  {model}");
Console.WriteLine($"Input:  {inputDir}");
Console.WriteLine($"Output: {outputDir}");
Console.WriteLine();

var pdfFiles = Directory.EnumerateFiles(inputDir, "*.pdf", SearchOption.AllDirectories)
    .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
    .ToList();

if (pdfFiles.Count == 0)
{
    Console.WriteLine("No PDF files found in the input folder. Nothing to do.");
    return 1;
}

var results = new List<StatementFile>();
var backendState = new string?[] { null }; // [0] = "files" | "vision" | "text" - learned, then reused


foreach (var pdf in pdfFiles)
{
    var name = Path.GetFileName(pdf);
    Console.WriteLine($"--- {name} ---");

    var record = new StatementFile { FileName = name };
    try
    {
        var json = await ClassifyWithLlm(http, model, pdf, name, maxPages, dpi, haveVision, haveText, backendState);
        var parsed = JsonSerializer.Deserialize<StatementFile>(json, jsonOptions);
        if (parsed is null || parsed.Statements is null)
            throw new Exception("LLM returned empty or malformed JSON.");

        record.Category = NormalizeCategory(parsed.Category, name);
        record.Institution = parsed.Institution;
        record.InstitutionType = parsed.InstitutionType;
        record.Statements = parsed.Statements;

        var targetFolder = Path.Combine(outputDir, record.Category);
        Directory.CreateDirectory(targetFolder);
        File.Copy(pdf, Path.Combine(targetFolder, name), overwrite: true);

        Console.WriteLine($"  -> {record.Category} / {record.Institution} " +
            $"({record.Statements.Count} statement(s)):");
        foreach (var s in record.Statements)
            Console.WriteLine($"     [{s.StatementType}] {s.AccountName}  " +
                            $"balance={s.Balance?.ToString("N2") ?? "?"}  " +
                            $"stmtDate={s.StatementDate:yyyy-MM-dd}  " +
                            $"due={s.DueDate?.ToString("yyyy-MM-dd") ?? "n/a"}");
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"  !! Failed: {ex.Message}");
        record.Error = true;
        record.ErrorMessage = ex.Message;
        record.Category = "Other";
        Directory.CreateDirectory(Path.Combine(outputDir, "Other"));
        File.Copy(pdf, Path.Combine(outputDir, "Other", name), overwrite: true);
    }

    results.Add(record);
}

var outputPath = Path.Combine(outputDir, "statements.json");
File.WriteAllText(outputPath, JsonSerializer.Serialize(results, jsonOptions));

Console.WriteLine();
Console.WriteLine($"Done. {results.Count} file(s) processed, " +
    $"{results.Count(r => !r.Error)} ok, {results.Count(r => r.Error)} failed.");
Console.WriteLine($"JSON written to: {outputPath}");
return 0;

// ---------------------------------------------------------------------------

static string BuildPrompt(string fileName, string source) =>
    "You are extracting data from financial statement documents. The file name is \"" + fileName + "\". " +
    "The document may contain MULTIPLE statements from the same or different institutions" +
    " (e.g. an annual report plus a monthly statement, or checking + credit card statements together)." +
    """

    Instructions:
    - Classify the document into exactly one category: "Fidelity", "Vanguard", "Bank", or "Other".
      "Bank" covers any bank, credit union, credit card, checking or savings statement that is
      not from Fidelity or Vanguard.
    - Extract EVERY distinct statement/account found in the document, one entry each.
    - Dates must be ISO format (yyyy-MM-dd). Use null for anything not present.
    - Balances: use the closing/current balance for "balance". Do not invent numbers.
    - statementDate = the date the statement was issued. dueDate = payment due date (credit cards /
      bank statements only; null for investment statements).

    Return ONLY a JSON object, no markdown, matching this shape exactly:
    {
      "category": "Fidelity" | "Vanguard" | "Bank" | "Other",
      "institution": "exact institution name as printed",
      "institutionType": "investment" | "bank" | "credit_card" | "unknown",
      "statements": [
        {
          "institution": "...",
          "accountName": "...",
          "accountNumber": "... or null",
          "statementType": "...",
          "openingBalance": 0.0,
          "closingBalance": 0.0,
          "balance": 0.0,
          "statementDate": "yyyy-MM-dd",
          "asOfDate": "yyyy-MM-dd or null",
          "dueDate": "yyyy-MM-dd or null",
          "amountDue": 0.0,
          "notes": ["..."]
        }
      ]
    }
    """ +
    source switch
    {
        "file" => "\nExtract the data from the attached PDF file.\n",
        "vision" => "\nExtract the data from the attached document page images (in page order).\n",
        _ => "\nExtract the data from the text below.\n",
    };

static async Task<string> ClassifyWithLlm(
    HttpClient http, string model, string pdfPath, string fileName,
    int maxPages, int dpi, bool canVision, bool canText, string?[] backendState)
{
    var order = new List<string> { "files" };
    if (canVision) order.Add("vision");
    if (canText) order.Add("text");
    if (backendState[0] is not null) order = new List<string> { backendState[0] };

    Exception? last = null;
    foreach (var mode in order)
    {
        try
        {
            var json = mode switch
            {
                "files" => await CallWithPdfFile(http, model, pdfPath, fileName),
                "vision" => await CallWithVision(http, model, pdfPath, fileName, maxPages, dpi),
                _ => await CallWithText(http, model, pdfPath, fileName),
            };
            backendState[0] = mode;
            return json;
        }
        catch (Exception ex)
        {
            last = ex;
            Console.Error.WriteLine($"  (backend {mode} failed: {Truncate(ex.Message, 160)} - trying next)");
        }
    }
    throw last!;
}

// Backend 1: OpenAI-style /files upload (proprietary; OpenAI + some gateways only).
static async Task<string> CallWithPdfFile(HttpClient http, string model, string pdfPath, string fileName)
{
    // 1. Upload the PDF so the model can read it (handles scanned pages too).
    string fileId;
    using (var form = new MultipartFormDataContent())
    {
        var fileBytes = await File.ReadAllBytesAsync(pdfPath);
        var content = new ByteArrayContent(fileBytes);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
        form.Add(content, "file", fileName);
        form.Add(new StringContent("assistant"), "purpose");
        var resp = await http.PostAsync("files", form);
        await EnsureSuccessAsync(resp, "file upload");
        var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
        fileId = doc.RootElement.GetProperty("id").GetString()
            ?? throw new Exception("file upload response missing 'id'");
    }

    // 2. Wait for the file to finish being checked (OpenAI quirk).
    for (var i = 0; i < 30; i++)
    {
        var statusResp = await http.GetAsync($"files/{fileId}");
        await EnsureSuccessAsync(statusResp, "file status");
        var statusDoc = JsonDocument.Parse(await statusResp.Content.ReadAsStringAsync());
        var status = statusDoc.RootElement.GetProperty("status").GetString();
        if (status is not ("uploaded" or "pending")) break;
        await Task.Delay(1000);
    }

    return await SendChatAsync(http, model, new object[]
    {
        new { type = "text", text = BuildPrompt(fileName, "file") },
        new { type = "file", filename = fileName, file_id = fileId },
    });
}

// Backend 2: standard OpenAI Vision API - rasterize pages, send base64 image parts.
// This is the vLLM / LM Studio path (any vision-capable OpenAI-compatible server).
static async Task<string> CallWithVision(HttpClient http, string model, string pdfPath, string fileName, int maxPages, int dpi)
{
    var pages = PdfTools.Rasterize(pdfPath, dpi, maxPages);
    if (pages.Count == 0) throw new Exception("no pages rasterized");

    var parts = new List<object>
    {
        new { type = "text", text = BuildPrompt(fileName, "vision") },
    };
    foreach (var (name, png) in pages)
    {
        var b64 = Convert.ToBase64String(png);
        parts.Add(new { type = "image_url", image_url = new { url = $"data:image/png;base64,{b64}" } });
    }
    return await SendChatAsync(http, model, parts.ToArray());
}

// Backend 3: plain text extraction via pdftotext (cheapest; fails on scanned PDFs).
static async Task<string> CallWithText(HttpClient http, string model, string pdfPath, string fileName)
{
    var text = PdfTools.ExtractText(pdfPath);
    if (text.Length > 120_000) text = text[..120_000];

    return await SendChatAsync(http, model, new object[]
    {
        new { type = "text", text = BuildPrompt(fileName, "text") + "\n\nPDF text:\n" + text },
    });
}

static async Task<string> SendChatAsync(HttpClient http, string model, object[] contentParts)
{
    var request = new
    {
        model,
        messages = new object[] { new { role = "user", content = contentParts } },
    };
    var payload = JsonSerializer.Serialize(request);
    var chatResp = await http.PostAsync("chat/completions",
        new StringContent(payload, Encoding.UTF8, "application/json"));
    await EnsureSuccessAsync(chatResp, "chat completion");
    var body = await chatResp.Content.ReadAsStringAsync();
    var chatDoc = JsonDocument.Parse(body);
    var raw = chatDoc.RootElement.GetProperty("choices")[0].GetProperty("message")
        .GetProperty("content").GetString() ?? string.Empty;

    // Models sometimes wrap JSON in code fences even when told not to.
    var start = raw.IndexOf('{');
    var end = raw.LastIndexOf('}');
    if (start < 0 || end <= start)
        throw new Exception("LLM response did not contain a JSON object.");
    return raw[start..(end + 1)];
}

static async Task EnsureSuccessAsync(HttpResponseMessage resp, string what)
{
    if (!resp.IsSuccessStatusCode)
    {
        var body = await resp.Content.ReadAsStringAsync();
        throw new Exception($"{what} failed: {(int)resp.StatusCode} {resp.ReasonPhrase} - {Truncate(body, 300)}");
    }
}

static string NormalizeCategory(string llmCategory, string fileName)
{
    var c = (llmCategory ?? string.Empty).Trim().ToLowerInvariant();
    if (c.Contains("fidel")) return "Fidelity";
    if (c.Contains("vanguard")) return "Vanguard";
    if (c.Contains("bank") || c.Contains("credit") || c.Contains("checking") || c.Contains("savings"))
        return "Bank";
    // fall back to filename hints
    var name = fileName.ToLowerInvariant();
    if (name.Contains("fidel")) return "Fidelity";
    if (name.Contains("vanguard")) return "Vanguard";
    if (name.Contains("bank") || name.Contains("card") || name.Contains("checking")) return "Bank";
    return "Other";
}

static string Truncate(string s, int n) => string.IsNullOrEmpty(s) ? "" : (s.Length <= n ? s : s[..n] + "...");
