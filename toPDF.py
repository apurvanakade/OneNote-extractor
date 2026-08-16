#!/usr/bin/env python3
"""
OneNote Notebook -> PDF Exporter (one PDF per section)
========================================================

Exports OneNote notebooks (via the Microsoft Graph API — works on any OS,
no desktop app needed) to PDF, one file per section, in most-recently-
modified-first order. Handwriting/ink is reconstructed from InkML stroke
data, since OneNote's plain content API silently drops it.

Full setup and usage instructions: see README.md.
Quick start: copy config.example.py to config.py, fill in CLIENT_ID, then
`python toPDF.py`.
"""

import re
import sys
import time
import base64
import unicodedata
import xml.etree.ElementTree as ET
from email import message_from_bytes
from pathlib import Path

import requests
import msal
from pypdf import PdfWriter, PdfReader
from playwright.sync_api import sync_playwright

try:
    from config import CLIENT_ID
except ImportError:
    sys.exit(
        "Missing config.py. Copy config.example.py to config.py and fill in "
        "your Azure app's client ID (see README.md 'One-time setup')."
    )

# ============================== CONFIG ======================================
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Notes.Read"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
OUTPUT_DIR = Path("onenote_export")
TOKEN_CACHE_FILE = Path(".onenote_token_cache.json")
PDF_PAPER_FORMAT = "Letter"  # change to "A4" if you prefer

# Microsoft doesn't publish a hard per-minute number for the OneNote API
# (it's absent from https://learn.microsoft.com/en-us/graph/throttling-limits)
# and OneNote's 429 responses don't include a Retry-After header, unlike most
# of Graph. So: pace every request conservatively up front, and back off
# with growing waits on any 429 we still hit.
MIN_REQUEST_INTERVAL_SEC = 1.1  # ~54 req/min, safely under the commonly-reported ~60/min cap
MAX_429_RETRIES = 6
INITIAL_BACKOFF_SEC = 5.0

# OneNote's plain content endpoint silently drops handwriting/drawings,
# replacing each stroke with `<!-- InkNode is not supported -->`. Fetching
# with includeInkML=true returns the strokes separately as InkML; their
# coordinates are HIMETRIC (1/100 mm) and share the same page-relative
# origin as the HTML's absolute-positioned elements, so this ratio converts
# them straight to the CSS px space OneNote's HTML already uses.
HIMETRIC_PER_PX = 2540 / 96
INK_NS = {"inkml": "http://www.w3.org/2003/InkML"}
XML_ID_ATTR = "{http://www.w3.org/XML/1998/namespace}id"
# =============================================================================


def get_access_token() -> str:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to start device login: {flow}")
        print("\n" + flow["message"] + "\n")
        result = app.acquire_token_by_device_flow(flow)  # blocks until you sign in

    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize())

    if "access_token" not in result:
        raise RuntimeError(f"Authentication failed: {result.get('error_description')}")

    return result["access_token"]


_last_request_time = 0.0


def graph_http_get(url: str, token: str, params=None) -> requests.Response:
    """Rate-limited, 429-retrying GET used for every Graph API call (JSON
    endpoints, page content, and embedded-image downloads alike)."""
    global _last_request_time
    headers = {"Authorization": f"Bearer {token}"}

    backoff = INITIAL_BACKOFF_SEC
    for attempt in range(MAX_429_RETRIES + 1):
        elapsed = time.monotonic() - _last_request_time
        if elapsed < MIN_REQUEST_INTERVAL_SEC:
            time.sleep(MIN_REQUEST_INTERVAL_SEC - elapsed)

        r = requests.get(url, headers=headers, params=params)
        _last_request_time = time.monotonic()

        if r.status_code != 429:
            return r
        if attempt == MAX_429_RETRIES:
            return r  # give up; caller's raise_for_status() will surface it

        retry_after = r.headers.get("Retry-After")
        wait = float(retry_after) if retry_after else backoff
        print(f"      (rate-limited (429), waiting {wait:.0f}s before retrying...)")
        time.sleep(wait)
        backoff *= 2

    return r  # unreachable


def graph_get(url: str, token: str, params=None) -> dict:
    r = graph_http_get(url, token, params=params)
    r.raise_for_status()
    return r.json()


def graph_get_all(url: str, token: str, params=None) -> list:
    """Follow @odata.nextLink pagination and return the combined 'value' list."""
    items = []
    while url:
        data = graph_get(url, token, params=params)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None  # nextLink already has query params baked in
    return items


def sanitize_filename(name: str) -> str:
    name = unicodedata.normalize("NFKD", name)
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    return name.strip()[:150] or "untitled"


def get_all_sections(token: str, notebook_id: str) -> list:
    """Recursively walk section groups and return
    [{'path': 'Group/Subgroup/SectionName', 'section': section_obj}, ...]."""
    results = []

    def walk_sections(sections_url, path_prefix):
        for sec in graph_get_all(sections_url, token):
            results.append({"path": path_prefix + sec["displayName"], "section": sec})

    def walk_section_groups(groups_url, path_prefix):
        for g in graph_get_all(groups_url, token):
            gid, gname = g["id"], g["displayName"]
            walk_sections(f"{GRAPH_ROOT}/me/onenote/sectionGroups/{gid}/sections",
                          f"{path_prefix}{gname}/")
            walk_section_groups(f"{GRAPH_ROOT}/me/onenote/sectionGroups/{gid}/sectionGroups",
                                 f"{path_prefix}{gname}/")

    walk_sections(f"{GRAPH_ROOT}/me/onenote/notebooks/{notebook_id}/sections", "")
    walk_section_groups(f"{GRAPH_ROOT}/me/onenote/notebooks/{notebook_id}/sectionGroups", "")

    # Most-recently-modified first, rather than whatever order the API
    # happened to return (which looks roughly alphabetical). ISO 8601
    # UTC timestamps sort correctly as plain strings.
    results.sort(key=lambda e: e["section"].get("lastModifiedDateTime", ""), reverse=True)
    return results


def fetch_page_content(token: str, page_id: str) -> tuple:
    """Return (html, inkml_bytes_or_None). See the includeInkML comment
    near the top of the file for why this isn't a plain GET."""
    url = f"{GRAPH_ROOT}/me/onenote/pages/{page_id}/content"
    r = graph_http_get(url, token, params={"includeInkML": "true"})
    r.raise_for_status()

    content_type = r.headers.get("Content-Type", "")
    if not content_type.startswith("multipart/"):
        return r.text, None

    # requests' r.text doesn't help for multipart bodies; parse the raw
    # bytes as a MIME message (re-adding the Content-Type header, which
    # lives in the HTTP response rather than the body itself).
    raw = f"Content-Type: {content_type}\r\n\r\n".encode() + r.content
    msg = message_from_bytes(raw)

    html_bytes, ink_bytes = None, None
    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_type()
        if ctype == "text/html":
            html_bytes = part.get_payload(decode=True)
        elif ctype == "application/inkml+xml":
            ink_bytes = part.get_payload(decode=True)

    if html_bytes is None:
        raise RuntimeError(f"No HTML part found for page {page_id}")
    return html_bytes.decode("utf-8"), ink_bytes


def render_ink_overlay(ink_xml: bytes) -> str:
    """Turn a page's InkML strokes into an absolutely-positioned inline SVG
    overlay, in the same page-relative pixel coordinates OneNote's HTML
    already uses, so it can be dropped straight into the page body."""
    root = ET.fromstring(ink_xml)

    contexts = {}
    for ctx in root.iterfind(".//inkml:context", INK_NS):
        channels = [c.get("name") for c in ctx.iterfind(".//inkml:channel", INK_NS)]
        contexts[ctx.get(XML_ID_ATTR)] = channels

    brushes = {}
    for brush in root.iterfind(".//inkml:brush", INK_NS):
        props = {p.get("name"): p.get("value") for p in brush.iterfind("inkml:brushProperty", INK_NS)}
        brushes[brush.get(XML_ID_ATTR)] = props

    def to_px(himetric: str) -> float:
        return float(himetric) / HIMETRIC_PER_PX

    paths = []
    max_x = max_y = 0.0
    for trace in root.iterfind(".//inkml:trace", INK_NS):
        props = brushes.get((trace.get("brushRef") or "").lstrip("#"), {})
        if props.get("rasterOp") == "noOperation":
            continue  # eraser / invisible stroke

        channels = contexts.get((trace.get("contextRef") or "").lstrip("#"), ["X", "Y"])
        if "X" not in channels or "Y" not in channels:
            continue
        xi, yi = channels.index("X"), channels.index("Y")

        points = []
        for point in (trace.text or "").strip().split(","):
            vals = point.split()
            if len(vals) <= max(xi, yi):
                continue
            x, y = to_px(vals[xi]), to_px(vals[yi])
            points.append((x, y))
            max_x, max_y = max(max_x, x), max(max_y, y)
        if len(points) < 2:
            continue

        stroke_width = max(to_px(props.get("width", "35")), to_px(props.get("height", "35")), 1.0)
        color = props.get("color", "#000000")
        opacity = 1.0 - float(props.get("transparency", "0"))

        d = "M " + " L ".join(f"{x:.1f} {y:.1f}" for x, y in points)
        paths.append(
            f'<path d="{d}" stroke="{color}" stroke-width="{stroke_width:.2f}" '
            f'stroke-opacity="{opacity:.3f}" fill="none" stroke-linecap="round" '
            f'stroke-linejoin="round" />'
        )

    if not paths:
        return ""

    margin = 40  # a little slack past the furthest stroke point
    return (
        f'<svg style="position:absolute;top:0;left:0;pointer-events:none" '
        f'width="{max_x + margin:.0f}" height="{max_y + margin:.0f}" overflow="visible">'
        + "".join(paths)
        + "</svg>"
    )


def inject_ink_overlay(html: str, svg: str) -> str:
    return re.sub(r"(<body[^>]*>)", lambda m: m.group(1) + svg, html, count=1)


def inline_resources(html: str, token: str) -> str:
    """OneNote page HTML references images via Graph resource URLs that
    require an auth header. Playwright can't send that header for
    page-internal requests, so we download each image and inline it as a
    base64 data URI instead."""
    def replace(match):
        url = match.group(1)
        if "graph.microsoft.com" not in url:
            return match.group(0)
        try:
            resp = graph_http_get(url, token)
            resp.raise_for_status()
            content_type = resp.headers.get("Content-Type", "image/png")
            b64 = base64.b64encode(resp.content).decode()
            return match.group(0).replace(url, f"data:{content_type};base64,{b64}")
        except Exception as e:
            print(f"      (warning: couldn't fetch an embedded image: {e})")
            return match.group(0)

    return re.sub(r'src="([^"]+)"', replace, html)


def render_html_to_pdf_bytes(browser, html: str) -> bytes:
    page = browser.new_page()
    page.set_content(html, wait_until="networkidle")
    pdf_bytes = page.pdf(
        format=PDF_PAPER_FORMAT,
        print_background=True,
        margin={"top": "0.4in", "bottom": "0.4in", "left": "0.4in", "right": "0.4in"},
    )
    page.close()
    return pdf_bytes


def export_section(token: str, browser, section: dict, section_path: str, out_dir: Path, filename_prefix: str = ""):
    pages = graph_get_all(f"{GRAPH_ROOT}/me/onenote/sections/{section['id']}/pages", token)
    if not pages:
        print("  (no pages, skipping)")
        return

    writer = PdfWriter()
    for i, pg in enumerate(pages, 1):
        title = pg.get("title") or f"Untitled page {i}"
        print(f"  [{i}/{len(pages)}] {title}")

        tmp_path = Path(f"_tmp_page_{i}.pdf")
        try:
            html, ink_xml = fetch_page_content(token, pg["id"])
            if ink_xml:
                svg = render_ink_overlay(ink_xml)
                if svg:
                    html = inject_ink_overlay(html, svg)
            html = inline_resources(html, token)
            pdf_bytes = render_html_to_pdf_bytes(browser, html)
            tmp_path.write_bytes(pdf_bytes)
            reader = PdfReader(str(tmp_path))
            start_index = len(writer.pages)
            for p in reader.pages:
                writer.add_page(p)
            writer.add_outline_item(title, start_index)
        except Exception as e:
            print(f"      !! failed to render this page, skipping: {e}")
            continue
        finally:
            tmp_path.unlink(missing_ok=True)

    if len(writer.pages) == 0:
        print("  (no pages rendered successfully, skipping)")
        return

    safe_name = sanitize_filename(section_path.replace("/", " - "))
    out_path = out_dir / f"{filename_prefix}{safe_name}.pdf"
    with open(out_path, "wb") as f:
        writer.write(f)
    print(f"  -> saved {out_path}")


def export_one_notebook(token: str, browser, notebook: dict):
    sections = get_all_sections(token, notebook["id"])
    print(f"Found {len(sections)} section(s).")

    out_dir = OUTPUT_DIR / sanitize_filename(notebook["displayName"])
    out_dir.mkdir(parents=True, exist_ok=True)

    index_width = len(str(len(sections)))
    for i, entry in enumerate(sections, 1):
        section, section_path = entry["section"], entry["path"]
        print(f"\n=== Section: {section_path} ===")
        try:
            prefix = f"{i:0{index_width}d} - "
            export_section(token, browser, section, section_path, out_dir, filename_prefix=prefix)
        except Exception as e:
            print(f"  !! failed to export section '{section_path}', skipping: {e}")


def export_notebook():
    if "PASTE-YOUR" in CLIENT_ID:
        sys.exit("Set CLIENT_ID in config.py first (see README.md 'One-time setup').")

    token = get_access_token()
    notebooks = graph_get_all(f"{GRAPH_ROOT}/me/onenote/notebooks", token)
    if not notebooks:
        print("No notebooks found on this account.")
        sys.exit(1)

    # One notebook per run, deliberately — even with the pacing/backoff in
    # graph_http_get, exporting multiple notebooks back-to-back still
    # reliably triggers 429 Too Many Requests. Run the script again for the
    # next notebook once this one finishes.
    print("\nAvailable notebooks:")
    for i, nb in enumerate(notebooks):
        print(f"  [{i}] {nb['displayName']}")
    idx = int(input("Pick a notebook number: ").strip())
    notebook = notebooks[idx]

    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        print(f"\n##### Notebook: {notebook['displayName']} #####")
        export_one_notebook(token, browser, notebook)
        browser.close()

    print("\nDone. Files are in:", OUTPUT_DIR.resolve())


if __name__ == "__main__":
    export_notebook()