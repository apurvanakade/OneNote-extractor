#!/usr/bin/env python3
"""
OneNote PDF -> LaTeX Transcriber
=================================

Walks onenote_export/ (as produced by toPDF.py) and produces one .tex file
per PDF, mirroring the same folder structure under onenote_export_latex/.

Most OneNote pages are handwritten ink baked into the PDF as vector strokes
(no embedded text, so pdftotext/OCR can't read them, especially not math
notation). So each page is rendered to a PNG and handed to the local Claude
Code CLI (`claude -p`, using your existing Claude subscription login rather
than a separate paid API key) with a prompt asking for a LaTeX transcription.
Any text pypdf can extract directly from the page is passed along as a hint
to ground the typed portions.

Requires: pypdf (pip install pypdf), poppler-utils' pdftoppm (already on
this machine), and the `claude` CLI logged in (this script reuses the same
binary this session runs from, via $CLAUDE_CODE_EXECPATH, or `claude` on
PATH).

Usage:
    python3 onenote_pdf_to_latex.py                     # prompt: pick a notebook under onenote_export/
    python3 onenote_pdf_to_latex.py path/to/one.pdf      # just one file
    python3 onenote_pdf_to_latex.py path/to/a/folder     # just one folder
    python3 onenote_pdf_to_latex.py --force              # ignore the up-to-date cache
    python3 onenote_pdf_to_latex.py --workers 6          # more parallel `claude` calls
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from pypdf import PdfReader

SOURCE_DIR = Path("onenote_export")
OUTPUT_DIR = Path("onenote_export_latex")
CLAUDE_TIMEOUT_SEC = 180

PROMPT_TEMPLATE = """You are transcribing one page of a handwritten/typed math notebook (exported from OneNote) into LaTeX source.

Text already extracted programmatically from this page (may be empty, partial, or out of order — use it to ground any typed portions, but the image is the source of truth):
---
{hint}
---

Read the attached image and transcribe ALL of its content into clean LaTeX:
- Use proper amsmath math mode (inline $...$, or pmatrix/bmatrix/align as appropriate) for every mathematical expression, matrix, or equation.
- Preserve structure: headings as \\textbf{{...}} or \\section*{{...}}, bullet lists as itemize, nesting/indentation, arrows as $\\to$ or $\\Rightarrow$.
- If something is a genuine hand-drawn diagram/sketch (not text or math), replace it with a short LaTeX comment describing it, e.g. % [diagram: sketch of ...] — do not try to redraw it.
- Ignore OneNote page furniture (timestamps, page titles already implied by the filename).
- Output ONLY the LaTeX body for this page: no \\documentclass, no \\begin{{document}}, no markdown code fences, no commentary.
- If the page is blank or has no meaningful content, output nothing at all.
"""

LATEX_PREAMBLE = r"""\documentclass[11pt]{article}
\usepackage{hyperref}
\usepackage{amsmath,amssymb,amsthm}
\usepackage[margin=1in]{geometry}
\usepackage{xcolor}
\usepackage{hyperref}
\begin{document}

"""
LATEX_POSTAMBLE = "\n\\end{document}\n"


def find_claude_binary() -> str:
    candidate = os.environ.get("CLAUDE_CODE_EXECPATH")
    if candidate and Path(candidate).exists():
        return candidate
    found = shutil.which("claude")
    if found:
        return found
    sys.exit("Couldn't find the `claude` CLI (checked $CLAUDE_CODE_EXECPATH and PATH). "
              "Make sure Claude Code is installed and you're logged in.")


def strip_code_fence(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines)
    return text.strip()


def transcribe_page(claude_bin: str, model: str, image_path: Path, hint_text: str) -> str:
    prompt = PROMPT_TEMPLATE.format(hint=hint_text.strip() or "(none — page is likely fully handwritten)")
    prompt += f"\n\nImage to transcribe: {image_path}"

    cmd = [
        claude_bin, "-p", prompt,
        "--add-dir", str(image_path.parent),
        "--tools", "Read",
        "--output-format", "text",
        "--model", model,
        "--no-session-persistence",
    ]
    for attempt in range(2):
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=CLAUDE_TIMEOUT_SEC)
        except subprocess.TimeoutExpired:
            if attempt == 0:
                continue
            return f"% [transcription timed out for {image_path.name}]"
        if result.returncode == 0 and result.stdout.strip():
            return strip_code_fence(result.stdout)
        if attempt == 0:
            continue
        return f"% [transcription failed for {image_path.name}: {result.stderr.strip()[:200]}]"


def render_pages(pdf_path: Path, tmp_dir: Path, dpi: int) -> list[Path]:
    prefix = tmp_dir / "page"
    subprocess.run(
        ["pdftoppm", "-png", "-r", str(dpi), str(pdf_path), str(prefix)],
        check=True, capture_output=True,
    )
    return sorted(tmp_dir.glob("page-*.png"))


def extract_text_per_page(pdf_path: Path) -> list[str]:
    reader = PdfReader(str(pdf_path))
    return [page.extract_text() or "" for page in reader.pages]


def process_pdf(pdf_path: Path, out_path: Path, claude_bin: str, model: str, dpi: int, workers: int):
    print(f"-> {pdf_path}")
    hints = extract_text_per_page(pdf_path)

    with tempfile.TemporaryDirectory(prefix="onenote_pages_") as tmp:
        images = render_pages(pdf_path, Path(tmp), dpi)
        if len(images) != len(hints):
            # Fall back to positional pairing as best-effort; pdftoppm and
            # pypdf should agree on page count, but don't hard-fail if not.
            hints = (hints + [""] * len(images))[: len(images)]

        results = [None] * len(images)

        def worker(i_image_hint):
            i, image, hint = i_image_hint
            print(f"   page {i + 1}/{len(images)}...")
            return i, transcribe_page(claude_bin, model, image, hint)

        with ThreadPoolExecutor(max_workers=workers) as pool:
            for i, body in pool.map(worker, [(i, img, hint) for i, (img, hint) in enumerate(zip(images, hints))]):
                results[i] = body

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w") as f:
        f.write(LATEX_PREAMBLE)
        for i, body in enumerate(results, 1):
            if not body.strip():
                continue
            f.write(f"% --- Page {i} ---\n{body.strip()}\n\n")
        f.write(LATEX_POSTAMBLE)
    print(f"   saved {out_path}")


def is_up_to_date(pdf_path: Path, out_path: Path) -> bool:
    return out_path.exists() and out_path.stat().st_mtime >= pdf_path.stat().st_mtime


def list_notebooks() -> list[Path]:
    if not SOURCE_DIR.is_dir():
        sys.exit(f"Source directory not found: {SOURCE_DIR}")
    return sorted(p for p in SOURCE_DIR.iterdir() if p.is_dir())


def prompt_for_notebook(notebooks: list[Path]) -> Path | None:
    print(f"Notebooks found in {SOURCE_DIR}/:\n")
    for i, nb in enumerate(notebooks, 1):
        print(f"  {i}. {nb.name}")
    print(f"  {len(notebooks) + 1}. (all notebooks)\n")

    choice = input(f"Which notebook to convert? [1-{len(notebooks) + 1}]: ").strip()
    try:
        index = int(choice)
    except ValueError:
        sys.exit(f"Not a number: {choice}")

    if index == len(notebooks) + 1:
        return None
    if 1 <= index <= len(notebooks):
        return notebooks[index - 1]
    sys.exit(f"Out of range: {choice}")


def collect_pdfs(paths: list[str]) -> list[Path]:
    if not paths:
        notebooks = list_notebooks()
        if not notebooks:
            sys.exit(f"No notebook folders found in {SOURCE_DIR}/.")
        chosen = prompt_for_notebook(notebooks)
        paths = [str(chosen)] if chosen else [str(SOURCE_DIR)]
    pdfs = []
    for p in paths:
        path = Path(p)
        if path.is_dir():
            pdfs.extend(sorted(path.rglob("*.pdf")))
        elif path.suffix.lower() == ".pdf":
            pdfs.append(path)
        else:
            sys.exit(f"Not a PDF or directory: {p}")
    return pdfs


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("paths", nargs="*", help="PDF file(s) or folder(s) to process (default: onenote_export/)")
    parser.add_argument("--force", action="store_true", help="Re-transcribe even if the .tex output is newer than the .pdf")
    parser.add_argument("--workers", type=int, default=3, help="Parallel `claude` calls per PDF (default: 3)")
    parser.add_argument("--model", default="sonnet", help="Model alias/name to pass to `claude` (default: sonnet)")
    parser.add_argument("--dpi", type=int, default=150, help="Page render resolution (default: 150)")
    args = parser.parse_args()

    claude_bin = find_claude_binary()
    pdfs = collect_pdfs(args.paths)
    if not pdfs:
        sys.exit("No PDFs found.")

    for pdf_path in pdfs:
        try:
            rel = pdf_path.relative_to(SOURCE_DIR)
        except ValueError:
            rel = Path(pdf_path.name)
        out_path = OUTPUT_DIR / rel.with_suffix(".tex")

        if not args.force and is_up_to_date(pdf_path, out_path):
            print(f"skip (up to date): {pdf_path}")
            continue

        try:
            process_pdf(pdf_path, out_path, claude_bin, args.model, args.dpi, args.workers)
        except Exception as e:
            print(f"   !! failed to process {pdf_path}, skipping: {e}")

    print("\nDone.")


if __name__ == "__main__":
    main()
