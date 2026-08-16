# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A two-stage personal pipeline that exports OneNote notebooks (handwriting included) to PDF, then transcribes each page to LaTeX:

1. **`toPDF.py`** — OneNote notebook -> PDF via the Microsoft Graph API, one PDF per section, output to `onenote_export/<Notebook Name>/<index> - <Section Path>.pdf`.
2. **`onenote_pdf_to_latex.py`** — walks `onenote_export/`, renders each PDF page to a PNG, and asks a local `claude -p` call to transcribe it into LaTeX, mirroring the folder structure under `onenote_export_latex/`.

The stages are independent and run separately; there is no shared code between them besides the `onenote_export/` directory convention.

## Commands

```bash
# One-time setup
cp config.example.py config.py   # then fill in CLIENT_ID — see README "One-time setup"
pip install -r requirements.txt
playwright install chromium

# Stage 1: export notebooks to PDF (interactive notebook picker, device-code login on first run)
python toPDF.py

# Stage 2: transcribe PDFs to LaTeX (interactive notebook picker if run with no args)
python onenote_pdf_to_latex.py
python onenote_pdf_to_latex.py path/to/one.pdf     # single file, no prompt
python onenote_pdf_to_latex.py path/to/a/folder    # single folder, no prompt
python onenote_pdf_to_latex.py --force             # re-transcribe even if .tex is newer than .pdf
python onenote_pdf_to_latex.py --workers 6         # more parallel `claude` calls

# Ad hoc diagnostics for a blank/failed page (dumps raw Graph responses to inspect)
python debug_inkml.py

# Sanity-check a script after edits (no test suite in this repo)
python3 -m py_compile toPDF.py onenote_pdf_to_latex.py
```

There is no test suite, linter, or build step configured for this project — `py_compile` is the only automated check available.

## Architecture notes

### `toPDF.py`

- Auth is MSAL device-code flow (`get_access_token`), with the resulting token cached in `.onenote_token_cache.json` (gitignored, holds a live refresh token) so subsequent runs don't need a fresh login.
- All Graph API calls go through `graph_http_get`, which self-paces requests (`MIN_REQUEST_INTERVAL_SEC`) and retries on 429 with backoff, because OneNote's endpoints don't publish a rate limit or send `Retry-After`. Every other Graph-calling function builds on top of this one — don't call `requests.get` directly against Graph.
- **Handwriting reconstruction** is the core trick: OneNote's plain content endpoint drops ink entirely, so pages are fetched with `includeInkML=true` (a multipart response, hand-parsed via `email.message_from_bytes` since `requests` doesn't do multipart). `render_ink_overlay` converts the returned InkML strokes (HIMETRIC coordinates) into an absolutely-positioned inline SVG using `HIMETRIC_PER_PX`, which is then injected into the page HTML's `<body>` before it's rendered to PDF via headless Chromium (Playwright).
- Embedded images in the page HTML are Graph resource URLs requiring an auth header that Playwright can't attach to page-internal requests, so `inline_resources` downloads and inlines each one as a base64 data URI before rendering.
- Sections are walked recursively through nested section groups (`get_all_sections`) and sorted most-recently-modified-first; each section's PDF filename is prefixed with a zero-padded index so a plain folder listing preserves that order.
- One notebook is exported per run by design (see the comment above the notebook picker in `export_notebook`) — exporting several back-to-back reliably triggers 429s even with the built-in pacing.
- Failures are caught and skipped at both page level (inside `export_section`'s per-page loop) and section level (`export_one_notebook`), so one bad page or section doesn't abort the whole notebook export.

### `onenote_pdf_to_latex.py`

- With no path arguments, prompts interactively for which notebook folder under `onenote_export/` to convert (`list_notebooks` / `prompt_for_notebook`); explicit file/folder arguments skip the prompt.
- Per PDF: `extract_text_per_page` (pypdf) pulls any programmatically-extractable text as a grounding hint, `render_pages` rasterizes each page to PNG via `pdftoppm`, then `transcribe_page` shells out to the `claude` CLI (`claude -p`, `--tools Read`, one call per page, parallelized across pages via `ThreadPoolExecutor`) with the image path and text hint embedded in the prompt.
- `find_claude_binary` resolves the `claude` binary via `$CLAUDE_CODE_EXECPATH` (set when this script is itself invoked from within a Claude Code session) before falling back to `PATH`.
- Output `.tex` files are skipped on re-run if newer than their source PDF (`is_up_to_date`), unless `--force` is passed.
- A failure transcribing one PDF is caught in `main()` and logged/skipped rather than aborting the rest of the batch.

## Gitignored / account-specific files

- `config.py` — Azure app client ID (copy from `config.example.py`)
- `.onenote_token_cache.json` — live OAuth refresh token
- `onenote_export/`, `onenote_export_latex/` — actual exported notebook content
- `debug_inkml.py`'s dump files (`_inkml_debug_response.bin`, `_plain_debug_response.html`) and `toPDF.py`'s per-page scratch files (`_tmp_page_*.pdf`)

None of these should ever be committed; `debug_inkml.py`'s dumps in particular can contain real handwritten page content.
