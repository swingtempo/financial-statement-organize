# Financial Statement Organizer

Takes a folder of PDF statements (Fidelity, Vanguard, bank), uses an LLM to read
each PDF (native PDF parsing, so scanned statements work too), classifies them,
copies them into per-institution folders, and writes a structured JSON summary.

## Setup

1. .NET 9 SDK (installed at `~/.dotnet`, already on PATH via `~/.bashrc`).
   Nothing else to install — PDF reading uses the bundled PDFium engine by
   default. (Optional: set `PDF_BACKEND=poppler` to prefer poppler-utils when
   they are on `PATH`; the app falls back to PDFium if poppler is missing.)
2. Put your key in the `.env` at the repo root (or set env vars, which win):

```
LLM_API_KEY=...                     # for vLLM/LM Studio any non-empty string works
LLM_BASE_URL=http://localhost:8000/v1   # vLLM, LM Studio, or OpenAI
LLM_MODEL=your-vision-model
INPUT_DIR=./input
OUTPUT_DIR=./organized
MAX_PAGES=20                           # vision backend: pages rasterized per PDF
```

> Backend ladder (first one that works is remembered for the rest of the run):
> 1. **files** – OpenAI's proprietary `/files` upload (real OpenAI only; *not a
>    standard* — vLLM / LM Studio don't have it)
> 2. **vision** – rasterizes the pages and sends base64 `image_url` parts in a
>    normal chat request (standard OpenAI Vision API — this is the **vLLM /
>    LM Studio** path; needs a vision-capable model, e.g. Qwen2.5-VL)
> 3. **text** – text extraction (cheapest, fails on scanned PDFs)
>
> PDF rasterization/extraction uses the bundled **PDFium** engine by default
> (`Patagames.Pdf` NuGet package — its native runtime ships inside the package
> for Windows / macOS / Linux, so there is nothing to install). Optionally set
> `PDF_BACKEND=poppler` in `.env` to use `pdftoppm` / `pdftotext` instead when
> they are on `PATH` (used only if available, else it falls back to managed).
> Both paths have been verified to produce identical results.

## Run

```bash
cd StatementOrganizer
dotnet run                                  # uses INPUT_DIR / OUTPUT_DIR from .env
dotnet run -- /path/to/pdfs /path/to/out   # or pass folders explicitly
```

## Output

```
organized/
  Fidelity/   <copied pdfs>
  Vanguard/   <copied pdfs>
  Bank/       <copied pdfs>
  Other/      <copied pdfs + anything that failed>
  statements.json
```

`statements.json` is an array — one entry per PDF, each containing the category,
institution, and **every statement found in the file** (files may contain
multiple statements). Per statement: institution, account name/number,
statement type, opening/closing balance, statement date, as-of date, due date,
amount due.
