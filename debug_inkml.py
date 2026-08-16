#!/usr/bin/env python3
"""One-off diagnostic: dump the raw includeInkML=true response for a single
OneNote page (plus the plain response for comparison), so we can see exactly
what the Graph API returns for ink before writing a real parser against it.

Run: python debug_inkml.py
Reuses the auth/token-cache machinery from toPDF.py, so it'll reuse your
existing login if you've already run that script once.
"""
from pathlib import Path

import requests

from toPDF import get_access_token, graph_get_all, GRAPH_ROOT


def pick(label, items, name_fn):
    print(f"\n{label}:")
    for i, item in enumerate(items):
        print(f"  [{i}] {name_fn(item)}")
    idx = input(f"Pick {label.lower()} number: ").strip()
    return items[int(idx)]


def main():
    token = get_access_token()

    notebooks = graph_get_all(f"{GRAPH_ROOT}/me/onenote/notebooks", token)
    nb = pick("Notebook", notebooks, lambda x: x["displayName"])

    sections = graph_get_all(f"{GRAPH_ROOT}/me/onenote/notebooks/{nb['id']}/sections", token)
    sec = pick("Section", sections, lambda x: x["displayName"])

    pages = graph_get_all(f"{GRAPH_ROOT}/me/onenote/sections/{sec['id']}/pages", token)
    pg = pick("Page (choose one with handwriting on it)", pages, lambda x: x.get("title") or "(untitled)")

    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_ROOT}/me/onenote/pages/{pg['id']}/content"

    r = requests.get(url, headers=headers, params={"includeInkML": "true"})
    print("\n[includeInkML=true]")
    print("Status:", r.status_code)
    print("Content-Type:", r.headers.get("Content-Type"))
    out = Path("_inkml_debug_response.bin")
    out.write_bytes(r.content)
    print("Wrote raw response ->", out.resolve())

    r2 = requests.get(url, headers=headers)
    print("\n[plain content, no includeInkML]")
    print("Status:", r2.status_code)
    Path("_plain_debug_response.html").write_text(r2.text)
    print("Wrote plain HTML ->", Path('_plain_debug_response.html').resolve())


if __name__ == "__main__":
    main()
