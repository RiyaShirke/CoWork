"""Standalone Kreuzberg test — feed a single PDF to the local Kreuzberg
service and dump the extracted Markdown so you can eyeball what the
downstream LLM will actually see.

Usage:
    python scripts\\test_kreuzberg.py "path\\to\\stock-statement.pdf"

The extracted content is printed to stdout AND saved to
state\\debug\\kreuzberg-test-<timestamp>.md so you can open it in any editor.
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")

KREUZBERG_URL = os.environ.get("KREUZBERG_URL", "http://localhost:8000").rstrip("/")
KREUZBERG_TIMEOUT = int(os.environ.get("KREUZBERG_TIMEOUT", "180"))
DEBUG_DIR = Path(os.environ.get("STATE_DIR", str(ROOT / "state"))) / "debug"


def health_check() -> bool:
    try:
        r = requests.get(f"{KREUZBERG_URL}/health", timeout=5)
        return r.status_code == 200
    except requests.RequestException as exc:
        print(f"[!] Kreuzberg unreachable at {KREUZBERG_URL}: {exc}")
        print("    Is the container running? Try: docker ps | findstr cowork-kreuzberg")
        return False


def extract(pdf_path: Path) -> dict | None:
    print(f"[*] Sending {pdf_path.name} ({pdf_path.stat().st_size:,} bytes) to {KREUZBERG_URL}/extract ...")
    with pdf_path.open("rb") as f:
        r = requests.post(
            f"{KREUZBERG_URL}/extract",
            files={"files": (pdf_path.name, f, "application/pdf")},
            timeout=KREUZBERG_TIMEOUT,
        )
    if r.status_code != 200:
        print(f"[!] HTTP {r.status_code}: {r.text[:500]}")
        return None
    results = r.json()
    if not isinstance(results, list) or not results:
        print(f"[!] Unexpected response shape: {results!r}")
        return None
    return results[0]


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print("Usage: python scripts\\test_kreuzberg.py <pdf-path>")
        return 2
    pdf_path = Path(argv[1]).resolve()
    if not pdf_path.exists():
        print(f"[!] File not found: {pdf_path}")
        return 2
    if not pdf_path.suffix.lower() == ".pdf":
        print(f"[!] Not a PDF: {pdf_path.suffix}")
        return 2

    if not health_check():
        return 1

    result = extract(pdf_path)
    if result is None:
        return 1

    content = (result.get("content") or "").strip()
    tables = result.get("tables") or []
    metadata = result.get("metadata") or {}

    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_md = DEBUG_DIR / f"kreuzberg-test-{stamp}__{pdf_path.stem}.md"
    out_json = DEBUG_DIR / f"kreuzberg-test-{stamp}__{pdf_path.stem}.full.json"
    out_md.write_text(content, encoding="utf-8")
    out_json.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")

    print()
    print("=" * 70)
    print(f"content: {len(content):,} chars")
    print(f"tables : {len(tables)} structured table(s)")
    print(f"metadata keys: {list(metadata.keys())}")
    print("=" * 70)
    print()
    preview = content if len(content) <= 4000 else content[:4000] + f"\n\n[... truncated, full {len(content):,} chars saved to {out_md.name}]"
    print(preview)
    print()
    print("=" * 70)
    print(f"Full Markdown : {out_md}")
    print(f"Full JSON     : {out_json}")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
