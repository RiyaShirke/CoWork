"""
Cowork worker — single-pass bill ingestion.

Picks the oldest unprocessed file from INBOX_DIR, OCRs it via PaddleOCR (Docker),
extracts structured rows via a local Ollama LLM, and appends them to an Excel
workbook. Idempotency is tracked by SHA256 hash in state/processed.json.

Intended to be triggered every 10-15 minutes by Windows Task Scheduler.
"""

from __future__ import annotations

import hashlib
import ast
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent
load_dotenv(ROOT / ".env")


def env_path(key: str, default: str) -> Path:
    return Path(os.environ.get(key, default)).resolve()


INBOX_DIR = env_path("INBOX_DIR", str(ROOT / "inbox"))
PROCESSED_DIR = env_path("PROCESSED_DIR", str(ROOT / "processed"))
FAILED_DIR = env_path("FAILED_DIR", str(ROOT / "failed"))
OUTPUT_DIR = env_path("OUTPUT_DIR", str(ROOT / "output"))
STATE_DIR = env_path("STATE_DIR", str(ROOT / "state"))
LOG_DIR = env_path("LOG_DIR", str(ROOT / "logs"))
OUTPUT_XLSX = env_path("OUTPUT_XLSX", str(OUTPUT_DIR / "bills.xlsx"))

BATCH_LIMIT = int(os.environ.get("BATCH_LIMIT", "1"))
ALLOWED_EXTS = tuple(
    e.strip().lower()
    for e in os.environ.get(
        "ALLOWED_EXTS", ".png,.jpg,.jpeg,.bmp,.tif,.tiff,.pdf"
    ).split(",")
    if e.strip()
)

PADDLE_IMAGE = os.environ.get("PADDLE_IMAGE", "paddlecloud/paddleocr:2.6-cpu-latest")
PADDLE_LANG = os.environ.get("PADDLE_LANG", "en")
PADDLE_CACHE_DIR = env_path("PADDLE_CACHE_DIR", str(STATE_DIR / "paddleocr-cache"))

OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://localhost:11434").rstrip("/")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.2:1b")
OLLAMA_TIMEOUT = int(os.environ.get("OLLAMA_TIMEOUT", "600"))
OLLAMA_NUM_CTX = int(os.environ.get("OLLAMA_NUM_CTX", "2048"))
OLLAMA_NUM_PREDICT = int(os.environ.get("OLLAMA_NUM_PREDICT", "1024"))
OLLAMA_KEEP_ALIVE = os.environ.get("OLLAMA_KEEP_ALIVE", "10m")
OLLAMA_MAX_OCR_CHARS = int(os.environ.get("OLLAMA_MAX_OCR_CHARS", "2000"))
OLLAMA_MAX_ROWS = int(os.environ.get("OLLAMA_MAX_ROWS", "12"))

STATE_FILE = STATE_DIR / "processed.json"
LOCK_FILE = STATE_DIR / "worker.lock"

EXCEL_COLUMNS = [
    "processed_at",
    "source_file",
    "vendor",
    "invoice_no",
    "invoice_date",
    "line_item",
    "quantity",
    "unit_price",
    "amount",
    "currency",
    "notes",
]


class OutputWorkbookLocked(RuntimeError):
    pass

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

for d in (
    INBOX_DIR,
    PROCESSED_DIR,
    FAILED_DIR,
    OUTPUT_DIR,
    STATE_DIR,
    LOG_DIR,
    PADDLE_CACHE_DIR,
):
    d.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "cowork.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("cowork")


# ---------------------------------------------------------------------------
# Lock (prevents overlapping scheduler runs)
# ---------------------------------------------------------------------------

LOCK_STALE_SECONDS = 60 * 60  # 1 hour


def acquire_lock() -> bool:
    if LOCK_FILE.exists():
        age = time.time() - LOCK_FILE.stat().st_mtime
        if age < LOCK_STALE_SECONDS:
            log.info("Lock present (age=%ds); another run is active. Exiting.", int(age))
            return False
        log.warning("Stale lock (age=%ds) — taking over.", int(age))
    LOCK_FILE.write_text(str(os.getpid()), encoding="utf-8")
    return True


def release_lock() -> None:
    try:
        LOCK_FILE.unlink(missing_ok=True)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Idempotency store
# ---------------------------------------------------------------------------


def load_state() -> dict:
    if not STATE_FILE.exists():
        return {}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        log.warning("processed.json is corrupt; starting empty.")
        return {}


def save_state(state: dict) -> None:
    tmp = STATE_FILE.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(state, indent=2), encoding="utf-8")
    tmp.replace(STATE_FILE)


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


# ---------------------------------------------------------------------------
# Inbox discovery
# ---------------------------------------------------------------------------


def discover_inbox() -> list[Path]:
    files = [
        p
        for p in INBOX_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in ALLOWED_EXTS
    ]
    files.sort(key=lambda p: p.stat().st_mtime)  # oldest first
    return files


# ---------------------------------------------------------------------------
# PaddleOCR via Docker
# ---------------------------------------------------------------------------

# Matches the `('TEXT', 0.98)` tuples PaddleOCR prints to stdout.
_OCR_TEXT_RE = re.compile(r"\('((?:[^'\\]|\\.)*)',\s*0?\.\d+\)")
_LAST_OCR_BOXES: list[dict] = []


def parse_ocr_boxes(stdout: str) -> list[dict]:
    boxes = []
    for line in stdout.splitlines():
        marker = "ppocr INFO: "
        if marker not in line:
            continue
        payload = line.split(marker, 1)[1].strip()
        if not payload.startswith("[[["):
            continue
        try:
            coords, text_info = ast.literal_eval(payload)
        except (SyntaxError, ValueError):
            continue
        if not isinstance(text_info, tuple) or not text_info:
            continue
        xs = [float(point[0]) for point in coords]
        ys = [float(point[1]) for point in coords]
        boxes.append({
            "text": str(text_info[0]).strip(),
            "x1": min(xs),
            "x2": max(xs),
            "y1": min(ys),
            "y2": max(ys),
            "cx": sum(xs) / len(xs),
            "cy": sum(ys) / len(ys),
        })
    return boxes


def run_paddleocr(image: Path) -> str:
    """Run PaddleOCR in a one-shot container; return joined text."""
    global _LAST_OCR_BOXES
    mount = image.parent.resolve()
    container_path = f"/data/{image.name}"
    cmd = [
        "docker", "run", "--rm",
        "-v", f"{mount}:/data",
        "-v", f"{PADDLE_CACHE_DIR.resolve()}:/root/.paddleocr",
        PADDLE_IMAGE,
        "python3", "-c", "import paddleocr; paddleocr.main()",
        "--image_dir", container_path,
        "--lang", PADDLE_LANG,
        "--use_gpu", "false",
    ]
    log.info("PaddleOCR: %s", " ".join(cmd))
    result = subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"PaddleOCR failed (exit {result.returncode}):\n{result.stderr[-2000:]}"
        )
    _LAST_OCR_BOXES = parse_ocr_boxes(result.stdout)
    matches = _OCR_TEXT_RE.findall(result.stdout)
    if not matches:
        # Fall back to raw stdout so the LLM still has *something* to work with.
        log.warning("No OCR tuples parsed; falling back to raw stdout.")
        return result.stdout.strip()
    return "\n".join(matches)


# ---------------------------------------------------------------------------
# Ollama extraction
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = (
    "You are an information-extraction engine. You receive raw OCR text from a "
    "single bill / invoice / receipt and must return STRICT JSON describing the "
    "line items.\n\n"
    "Return an object with one key 'rows' whose value is a JSON array. Each row "
    "is an object with these fields (use empty string '' or null if not present):\n"
    "  vendor, invoice_no, invoice_date, line_item, quantity, unit_price, "
    "amount, currency, notes\n\n"
    "Rules:\n"
    "- One element of 'rows' per important line item on the bill.\n"
    f"- Return at most {OLLAMA_MAX_ROWS} rows. If there are more items, combine "
    "the remaining items into one summary row.\n"
    "- If the bill has no itemized lines, return a single row summarising the bill.\n"
    "- Dates as ISO YYYY-MM-DD when possible, else keep the original string.\n"
    "- Numbers as plain numbers (no currency symbol inside quantity/unit_price/amount).\n"
    "- Currency as ISO code (USD, INR, EUR...) when identifiable; else ''.\n"
    "- Keep line_item and notes short. Use at most 80 characters each.\n"
    "- Never invent data not visible in the OCR text.\n"
    "- Output JSON only. No markdown, no commentary."
)


_IMPORTANT_OCR_RE = re.compile(
    r"invoice|bill|receipt|date|gst|total|subtotal|tax|amount|qty|quantity|"
    r"price|rate|item|description|vendor|paid|balance|currency|rs\.?|inr|usd|\d",
    re.IGNORECASE,
)


def compact_ocr_text(ocr_text: str) -> str:
    """Keep the most useful OCR lines when a long image would slow the LLM."""
    lines = [line.strip() for line in ocr_text.splitlines() if line.strip()]
    if not lines:
        return ocr_text.strip()

    deduped = list(dict.fromkeys(lines))
    compact = "\n".join(deduped)
    if len(compact) <= OLLAMA_MAX_OCR_CHARS:
        return compact

    important = [line for line in deduped if _IMPORTANT_OCR_RE.search(line)]
    source = important if important else deduped
    compact = "\n".join(source)
    if len(compact) <= OLLAMA_MAX_OCR_CHARS:
        return compact

    head_chars = int(OLLAMA_MAX_OCR_CHARS * 0.65)
    tail_chars = OLLAMA_MAX_OCR_CHARS - head_chars
    return compact[:head_chars].rstrip() + "\n...\n" + compact[-tail_chars:].lstrip()


def call_ollama(ocr_text: str) -> list[dict]:
    prompt_ocr_text = compact_ocr_text(ocr_text)
    if len(prompt_ocr_text) != len(ocr_text):
        log.info(
            "Ollama: compacted OCR from %d to %d chars",
            len(ocr_text),
            len(prompt_ocr_text),
        )
    payload = {
        "model": OLLAMA_MODEL,
        "prompt": (
            f"{SYSTEM_PROMPT}\n\n"
            f"OCR_TEXT:\n<<<\n{prompt_ocr_text}\n>>>"
        ),
        "stream": False,
        "format": "json",
        "keep_alive": OLLAMA_KEEP_ALIVE,
        "options": {
            "temperature": 0.1,
            "num_ctx": OLLAMA_NUM_CTX,
            "num_predict": OLLAMA_NUM_PREDICT,
        },
    }
    log.info(
        "Ollama: POST /api/generate (model=%s, chars=%d, timeout=%ds)",
        OLLAMA_MODEL,
        len(prompt_ocr_text),
        OLLAMA_TIMEOUT,
    )
    r = requests.post(f"{OLLAMA_URL}/api/generate", json=payload, timeout=OLLAMA_TIMEOUT)
    r.raise_for_status()
    body = r.json()
    raw = body.get("response", "").strip()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        # Some models still wrap JSON in prose; grab the outermost {...}.
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if not m:
            raise RuntimeError(f"Ollama response was not JSON: {raw[:500]}")
        data = json.loads(m.group(0))
    rows = data.get("rows", [])
    if not isinstance(rows, list):
        raise RuntimeError(f"'rows' is not a list: {data!r}")
    return rows


# ---------------------------------------------------------------------------
# Fallback extraction
# ---------------------------------------------------------------------------

_AMOUNT_RE = re.compile(r"(?:rs\.?|inr|\u20b9)?\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)", re.IGNORECASE)
_DATE_RE = re.compile(
    r"\b(\d{1,2}[-/ ](?:\d{1,2}|[A-Za-z]{3,9})[-/ ]\d{2,4}|"
    r"\d{4}[-/]\d{1,2}[-/]\d{1,2})\b"
)
_INVOICE_RE = re.compile(r"\b(?:invoice|bill|receipt)\s*(?:no\.?|#|number)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-\/]+)", re.IGNORECASE)


def clean_amount(value: str) -> str:
    value = value.strip().replace(",", "")
    if value.count(".") > 1:
        left, right = value.rsplit(".", 1)
        value = left.replace(".", "") + "." + right
    return value


def box_text_near(boxes: list[dict], y: float, x_min: float, x_max: float) -> str:
    candidates = [
        box for box in boxes
        if x_min <= box["cx"] <= x_max and abs(box["cy"] - y) <= 12
    ]
    candidates.sort(key=lambda box: box["x1"])
    return " ".join(box["text"] for box in candidates).strip()


def invoice_meta_from_ocr(lines: list[str]) -> tuple[str, str, str, str]:
    vendor = next((line for line in lines[:8] if not re.search(r"invoice|bill|receipt|store|order", line, re.I)), lines[0])
    joined = "\n".join(lines)
    invoice_no = ""
    invoice_match = _INVOICE_RE.search(joined)
    if invoice_match:
        invoice_no = invoice_match.group(1)
    date_match = _DATE_RE.search(joined)
    invoice_date = date_match.group(1) if date_match else ""
    currency = "INR" if re.search(r"\b(?:inr|rs\.?)\b|\u20b9|rupees", joined, re.I) else ""
    return vendor, invoice_no, invoice_date, currency


def table_fallback_rows(lines: list[str]) -> list[dict]:
    boxes = _LAST_OCR_BOXES
    if not boxes:
        return []

    header_candidates = [
        box for box in boxes
        if re.search(r"name.*product|product.*service|description|hsn|rate|taxable", box["text"], re.I)
    ]
    if not header_candidates:
        return []
    header_y = min(box["cy"] for box in header_candidates)
    total_candidates = [
        box for box in boxes
        if box["cy"] > header_y + 30 and box["cx"] < 560 and re.fullmatch(r"total", box["text"], re.I)
    ]
    total_y = min((box["cy"] for box in total_candidates), default=header_y + 430)
    table_boxes = [box for box in boxes if header_y + 15 <= box["cy"] <= total_y - 8]

    product_boxes = [
        box for box in table_boxes
        if 170 <= box["cx"] <= 415 and not re.fullmatch(r"\d+", box["text"])
    ]
    product_boxes.sort(key=lambda box: (box["cy"], box["x1"]))

    vendor, invoice_no, invoice_date, currency = invoice_meta_from_ocr(lines)
    rows: list[dict] = []
    current: dict | None = None

    for box in product_boxes:
        y = box["cy"]
        total = box_text_near(table_boxes, y, 930, 1020)
        rate = box_text_near(table_boxes, y, 600, 685)
        qty = box_text_near(table_boxes, y, 515, 590)
        taxable = box_text_near(table_boxes, y, 700, 790)

        if total and rate:
            current = {
                "vendor": vendor,
                "invoice_no": invoice_no,
                "invoice_date": invoice_date,
                "line_item": box["text"],
                "quantity": qty,
                "unit_price": clean_amount(rate),
                "amount": clean_amount(total),
                "currency": currency,
                "notes": f"Taxable value {clean_amount(taxable)}" if taxable else "",
            }
            rows.append(current)
        elif current and box["cy"] - rows[-1].get("_last_y", y) < 55:
            current["line_item"] = f"{current['line_item']} {box['text']}".strip()

        if current:
            current["_last_y"] = y
        if len(rows) >= OLLAMA_MAX_ROWS:
            break

    for row in rows:
        row.pop("_last_y", None)
    return rows


def fallback_extract_rows(ocr_text: str) -> list[dict]:
    """Small deterministic backup when the LLM returns malformed JSON."""
    lines = [line.strip() for line in ocr_text.splitlines() if line.strip()]
    if not lines:
        return [{"notes": "No OCR text available"}]

    table_rows = table_fallback_rows(lines)
    if table_rows:
        return table_rows

    vendor, invoice_no, invoice_date, currency = invoice_meta_from_ocr(lines)

    rows = []
    for line in lines:
        if re.search(r"subtotal|sub total|total|tax|cgst|sgst|igst|amount|balance", line, re.I):
            continue
        amount_matches = _AMOUNT_RE.findall(line)
        if not amount_matches:
            continue
        amount = amount_matches[-1].replace(",", "")
        if len(amount) > 10:
            continue
        item = _AMOUNT_RE.sub("", line).strip(" :-\t")
        if not item or item.isdigit():
            continue
        rows.append({
            "vendor": vendor,
            "invoice_no": invoice_no,
            "invoice_date": invoice_date,
            "line_item": item[:120],
            "quantity": "",
            "unit_price": "",
            "amount": amount,
            "currency": currency,
            "notes": "Fallback OCR extraction",
        })
        if len(rows) >= OLLAMA_MAX_ROWS:
            break

    if rows:
        return rows
    total_lines = [line for line in lines if re.search(r"total|amount|balance", line, re.I)]
    note = " | ".join(total_lines[-3:] or lines[:5])
    return [{
        "vendor": vendor,
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "line_item": "Bill summary",
        "quantity": "",
        "unit_price": "",
        "amount": "",
        "currency": currency,
        "notes": note[:200],
    }]


# ---------------------------------------------------------------------------
# Excel append
# ---------------------------------------------------------------------------


def append_rows(rows: list[dict], source_file: str) -> int:
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    if OUTPUT_XLSX.exists():
        wb = load_workbook(OUTPUT_XLSX)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "bills"
        ws.append(EXCEL_COLUMNS)

    now_iso = datetime.now().isoformat(timespec="seconds")
    written = 0
    for row in rows:
        ws.append([
            now_iso,
            source_file,
            row.get("vendor", "") or "",
            row.get("invoice_no", "") or "",
            row.get("invoice_date", "") or "",
            row.get("line_item", "") or "",
            row.get("quantity", "") if row.get("quantity") is not None else "",
            row.get("unit_price", "") if row.get("unit_price") is not None else "",
            row.get("amount", "") if row.get("amount") is not None else "",
            row.get("currency", "") or "",
            row.get("notes", "") or "",
        ])
        written += 1

    try:
        wb.save(OUTPUT_XLSX)
    except PermissionError as exc:
        raise OutputWorkbookLocked(
            f"{OUTPUT_XLSX} is locked. Close Excel and run the worker again."
        ) from exc
    return written


def ensure_output_writable() -> None:
    if not OUTPUT_XLSX.exists():
        return
    try:
        with OUTPUT_XLSX.open("a+b"):
            pass
    except PermissionError as exc:
        raise OutputWorkbookLocked(
            f"{OUTPUT_XLSX} is locked. Close Excel and run the worker again."
        ) from exc


# ---------------------------------------------------------------------------
# File movement
# ---------------------------------------------------------------------------


def move_to(path: Path, dest_dir: Path) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = dest_dir / f"{stamp}__{path.name}"
    shutil.move(str(path), str(dest))
    return dest


# ---------------------------------------------------------------------------
# Per-file pipeline
# ---------------------------------------------------------------------------


def process_one(path: Path, state: dict) -> None:
    log.info("---- processing %s ----", path.name)
    digest = sha256_of(path)

    prior = state.get(digest)
    if prior and prior.get("status") == "success":
        log.info("Already processed (hash=%s, on %s). Moving to processed/.",
                 digest[:10], prior.get("processed_at"))
        try:
            move_to(path, PROCESSED_DIR)
        except OSError as e:
            log.warning("Could not move duplicate: %s", e)
        return

    try:
        ensure_output_writable()
        ocr_text = run_paddleocr(path)
        if not ocr_text.strip():
            raise RuntimeError("Empty OCR result")
        log.info("OCR: %d chars extracted", len(ocr_text))

        try:
            rows = call_ollama(ocr_text)
            log.info("LLM: %d row(s) extracted", len(rows))
        except Exception as exc:
            log.warning("LLM extraction failed; using fallback OCR extraction: %s", exc)
            rows = fallback_extract_rows(ocr_text)
            log.info("Fallback: %d row(s) extracted", len(rows))

        if not rows:
            rows = [{"notes": "LLM returned zero rows", "line_item": ocr_text[:200]}]

        n = append_rows(rows, source_file=path.name)
        dest = move_to(path, PROCESSED_DIR)

        state[digest] = {
            "source_file": path.name,
            "processed_at": datetime.now().isoformat(timespec="seconds"),
            "status": "success",
            "rows_added": n,
            "moved_to": str(dest),
        }
        save_state(state)
        log.info("OK: appended %d row(s) to %s", n, OUTPUT_XLSX)

    except OutputWorkbookLocked as exc:
        log.error("SKIPPED: %s", exc)

    except Exception as exc:
        log.exception("FAILED: %s", exc)
        try:
            dest = move_to(path, FAILED_DIR)
        except OSError:
            dest = path
        state[digest] = {
            "source_file": path.name,
            "processed_at": datetime.now().isoformat(timespec="seconds"),
            "status": "failed",
            "error": str(exc),
            "moved_to": str(dest),
        }
        save_state(state)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    log.info("Cowork worker run starting. INBOX=%s", INBOX_DIR)

    if not acquire_lock():
        return 0

    try:
        files = discover_inbox()
        log.info("Inbox: %d candidate file(s)", len(files))
        if not files:
            return 0

        if BATCH_LIMIT > 0:
            files = files[:BATCH_LIMIT]

        state = load_state()
        for f in files:
            process_one(f, state)

        return 0
    finally:
        release_lock()
        log.info("Cowork worker run finished.")


if __name__ == "__main__":
    raise SystemExit(main())
