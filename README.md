# OneNote -> PDF -> LaTeX

Exports OneNote notebooks to PDF (handwriting included) via the Microsoft
Graph API, then optionally transcribes each page to LaTeX using Claude.

Two independent stages, run in order:

1. **`toPDF.py`** — OneNote notebook(s) -> PDF, one file per section.
2. **`onenote_pdf_to_latex.py`** — those PDFs -> one `.tex` file per PDF.

## How it works

- Talks to the Microsoft Graph API directly, so it runs on any OS — no
  OneNote desktop app required.
- **Handwriting**: OneNote's plain content endpoint silently drops ink,
  replacing every stroke with an HTML comment. This script instead requests
  `includeInkML=true`, parses the returned InkML stroke data (coordinates,
  color, width, opacity), and re-renders it as an SVG overlay positioned to
  match OneNote's own page coordinates.
- **Section order**: sections are exported most-recently-modified first
  (across nested section groups too), and each section's PDF filename is
  prefixed with a zero-padded index so a plain folder listing preserves
  that order.
- **Rate limiting**: Microsoft doesn't publish a hard request-rate limit for
  the OneNote API, and its 429 responses don't include a `Retry-After`
  header (unlike most of Graph). Every request is self-paced to a
  conservative rate, with exponential backoff if a 429 still comes through.
- **Resilience**: a failure on one page, section, or notebook is logged and
  skipped rather than aborting the whole run — useful when exporting many
  notebooks unattended.

## One-time setup

### 1. Register an Azure AD app

This is how you get permission to read your own notebooks via the API —
about 3 minutes, no cost.

1. Go to [portal.azure.com](https://portal.azure.com) -> **App
   registrations** -> **New registration**.
   - Name: anything, e.g. "OneNote PDF Exporter"
   - Supported account types: **Accounts in any organizational directory
     and personal Microsoft accounts**
   - Redirect URI: leave blank for now
   - Click **Register**.
2. Copy the **Application (client) ID** from the Overview page.
3. Sidebar -> **Authentication** -> **Add a platform** -> **Mobile and
   desktop applications** -> check the suggested
   `https://login.microsoftonline.com/common/oauth2/nativeclient` redirect
   URI -> Configure. Under **Advanced settings**, set **Allow public client
   flows** to Yes -> Save.
4. Sidebar -> **API permissions** -> **Add a permission** -> **Microsoft
   Graph** -> **Delegated permissions** -> search "Notes" -> check
   **Notes.Read** -> Add permissions.
   - If this is a work/school account and you see "admin consent
     required", ask your Microsoft 365 admin to grant it — or sign in with
     a personal Microsoft account instead, which doesn't need that.

### 2. Configure this project

```bash
cp config.example.py config.py
```

Open `config.py` and paste in the client ID from step 1. `config.py` is
gitignored — it's the one file here specific to your Microsoft account.

### 3. Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

`onenote_pdf_to_latex.py` additionally needs:
- `pdftoppm`, from `poppler-utils` (`apt install poppler-utils` /
  `brew install poppler`)
- the [Claude Code](https://claude.com/claude-code) CLI, logged in — it
  reuses your existing subscription rather than a separate API key

## Usage

### Export notebooks to PDF

```bash
python toPDF.py
```

First run opens a device-code login: it prints a URL and a short code, you
enter it in any browser and sign in. A token cache (`.onenote_token_cache.json`,
also gitignored — it holds a live refresh token) is saved locally so you
won't have to log in again.

You'll then see a numbered list of your notebooks:

```
Available notebooks:
  [0] 2026 Spring Discrete Math
  [1] Graph Theory
  ...
Pick a notebook number:
```

Enter a number to export that notebook. One notebook per run, deliberately
— exporting several back-to-back reliably triggers 429s even with the
built-in pacing/backoff, so run the script again for the next one.

Output: `onenote_export/<Notebook Name>/<index> - <Section Path>.pdf`

### Transcribe to LaTeX

```bash
python onenote_pdf_to_latex.py
```

With no arguments, it lists the notebook folders under `onenote_export/`
and asks which one to convert:

```
Notebooks found in onenote_export/:

  1. 2026 Spring Computational Math
  2. 2026 Spring Discrete Math
  3. 2026 Spring Monte Carlo Methods
  4. (all notebooks)

Which notebook to convert? [1-4]:
```

For the chosen notebook (or all of them), it renders each PDF page to an
image and asks Claude to transcribe it (math notation included) into LaTeX.
Output mirrors the same folder structure under `onenote_export_latex/`.
Already-transcribed files are skipped on re-runs unless the source PDF is
newer (or you pass `--force`). A failure on one PDF is logged and skipped
rather than aborting the rest of the run.

```bash
python onenote_pdf_to_latex.py path/to/one.pdf     # just one file, no prompt
python onenote_pdf_to_latex.py path/to/a/folder    # just one folder, no prompt
python onenote_pdf_to_latex.py --workers 6         # more parallel Claude calls
```

## Troubleshooting

**A page's PDF came out blank.** Almost always means it's mostly/entirely
handwritten and the ink overlay failed to parse for that page. Run
`debug_inkml.py`, pick the affected page, and inspect the two files it
writes (raw includeInkML response and plain HTML) to see what came back.

**429 Too Many Requests.** The script already paces and retries
automatically; if you still hit persistent 429s, raise
`MIN_REQUEST_INTERVAL_SEC` near the top of `toPDF.py`.

## What's gitignored, and why

- `config.py` — your Azure app client ID
- `.onenote_token_cache.json` — a live OAuth refresh token
- `onenote_export/`, `onenote_export_latex/` — your actual notebook content
- `debug_inkml.py`'s dump files — can contain real page content
