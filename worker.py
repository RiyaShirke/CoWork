"""
Cowork worker — atomic stock-statement ingestion.

Pulls the oldest unprocessed file from INBOX_DIR, extracts text (pdfplumber for
text PDFs, PaddleOCR via Docker for images and scanned PDFs), asks a local
Ollama LLM to map it to a strict JSON schema (report metadata + product rows +
optional batch rows), and appends the rows to a single Excel workbook.

Pipeline guarantees:
  * Per-file work is all-or-nothing. The workbook, the state file, and the
    file's location in inbox/processed/failed are either ALL updated together
    or NOT AT ALL. A crash mid-run leaves nothing half-done.
  * File-specific failures (corrupt PDF, malformed LLM output) increment an
    attempt counter; after RETRY_LIMIT attempts the file moves to failed/.
  * Environment failures (Docker down, Ollama unreachable, workbook open in
    Excel) do NOT burn attempts — they just skip the run.
  * Idempotency: same content hash never produces duplicate rows.

Intended to be triggered every 10-15 minutes by Windows Task Scheduler.
"""

from __future__ import annotations

import hashlib
import io
import json
import logging
import logging.handlers
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parent
load_dotenv(ROOT / ".env")


def _env_path(key: str, default: Path) -> Path:
    return Path(os.environ.get(key, str(default))).resolve()


def _env_int(key: str, default: int) -> int:
    try:
        return int(os.environ.get(key, str(default)))
    except ValueError:
        return default


INBOX_DIR = _env_path("INBOX_DIR", ROOT / "inbox")
PROCESSED_DIR = _env_path("PROCESSED_DIR", ROOT / "processed")
FAILED_DIR = _env_path("FAILED_DIR", ROOT / "failed")
OUTPUT_DIR = _env_path("OUTPUT_DIR", ROOT / "output")
STATE_DIR = _env_path("STATE_DIR", ROOT / "state")
LOG_DIR = _env_path("LOG_DIR", ROOT / "logs")
OUTPUT_XLSX = _env_path("OUTPUT_XLSX", OUTPUT_DIR / "bills.xlsx")

BATCH_LIMIT = _env_int("BATCH_LIMIT", 1)
RETRY_LIMIT = _env_int("RETRY_LIMIT", 3)
SUBPROCESS_TIMEOUT = _env_int("SUBPROCESS_TIMEOUT", 600)
ALLOWED_EXTS = tuple(
    e.strip().lower()
    for e in os.environ.get(
        "ALLOWED_EXTS", ".pdf,.png,.jpg,.jpeg,.bmp,.tif,.tiff"
    ).split(",")
    if e.strip()
)

PADDLE_IMAGE = os.environ.get("PADDLE_IMAGE", "paddlecloud/paddleocr:2.6-cpu-latest")
PADDLE_LANG = os.environ.get("PADDLE_LANG", "en")
PADDLE_CACHE_DIR = _env_path("PADDLE_CACHE_DIR", STATE_DIR / "paddleocr-cache")
PDF_OCR_DPI = _env_int("PDF_OCR_DPI", 200)

OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://localhost:11434").rstrip("/")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.1:8b")
OLLAMA_TIMEOUT = _env_int("OLLAMA_TIMEOUT", 600)
OLLAMA_NUM_CTX = _env_int("OLLAMA_NUM_CTX", 8192)
OLLAMA_NUM_PREDICT = _env_int("OLLAMA_NUM_PREDICT", 16384)
OLLAMA_KEEP_ALIVE = os.environ.get("OLLAMA_KEEP_ALIVE", "10m")
OLLAMA_MAX_TEXT_CHARS = _env_int("OLLAMA_MAX_TEXT_CHARS", 24000)
OLLAMA_RETRY_ATTEMPTS = _env_int("OLLAMA_RETRY_ATTEMPTS", 3)
OLLAMA_RETRY_BACKOFF_SEC = _env_int("OLLAMA_RETRY_BACKOFF_SEC", 5)

# Minimum chars of native PDF text below which we treat the PDF as effectively
# image-based and run OCR as well. Stock statements with thousands of product
# rows but only a tiny embedded text layer (header only) fall here.
PDF_TEXT_MIN_CHARS = _env_int("PDF_TEXT_MIN_CHARS", 1500)

NOTIFY_WEBHOOK_URL = os.environ.get("NOTIFY_WEBHOOK_URL", "").strip()

STATE_FILE = STATE_DIR / "processed.json"
LOCK_FILE = STATE_DIR / "worker.lock"
DEBUG_DIR = STATE_DIR / "debug"
LOCK_STALE_SECONDS = _env_int("LOCK_STALE_SECONDS", 60 * 60)

REPORT_META_COLUMNS = [
    "DistributorName",
    "DivisionName",
    "ReportFromDate",
    "ReportToDate",
    "TotalStockValue",
    "TotalSalesValue",
]

PRODUCT_COLUMNS = [
    "ProductCode",
    "ProductName",
    "Packing",
    "OpeningStock",
    "PurchaseQty",
    "SalesQty",
    "ClosingStock",
    "StockValue",
    "SalesValue",
    "AdjustmentQty",
    "PreviousMonthStock",
    "TwoMonthOldStock",
    "OldStock120Days",
    "ExpiryWithin3Months",
]

BATCH_COLUMNS = [
    "BatchNumber",
    "ExpiryDate",
    "BatchStock",
    "MRP",
    "BatchValue",
]

EXCEL_COLUMNS = (
    ["processed_at", "source_file"]
    + REPORT_META_COLUMNS
    + PRODUCT_COLUMNS
    + BATCH_COLUMNS
)

PDF_EXTS = {".pdf"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


class FileExtractionError(RuntimeError):
    """File is bad in a way that re-running won't help. Counts against attempts."""


class EnvironmentSkip(RuntimeError):
    """Worker can't run right now (Docker/Ollama down, workbook locked). Does NOT count."""


# ---------------------------------------------------------------------------
# Bootstrap dirs + logging
# ---------------------------------------------------------------------------

for _d in (
    INBOX_DIR,
    PROCESSED_DIR,
    FAILED_DIR,
    OUTPUT_DIR,
    STATE_DIR,
    LOG_DIR,
    PADDLE_CACHE_DIR,
):
    _d.mkdir(parents=True, exist_ok=True)


def _setup_logging() -> logging.Logger:
    logger = logging.getLogger("cowork")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    rot = logging.handlers.RotatingFileHandler(
        LOG_DIR / "cowork.log",
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    rot.setFormatter(fmt)
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(fmt)
    logger.addHandler(rot)
    logger.addHandler(stream)
    return logger


log = _setup_logging()


# ---------------------------------------------------------------------------
# Lock (PID-aware so a crashed worker doesn't block the scheduler for an hour)
# ---------------------------------------------------------------------------


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        # Process exists, we just can't signal it.
        return True
    except OSError:
        return False
    return True


def acquire_lock() -> bool:
    if LOCK_FILE.exists():
        try:
            payload = json.loads(LOCK_FILE.read_text(encoding="utf-8"))
            pid = int(payload.get("pid", -1))
            started_at = payload.get("started_at", "")
        except (ValueError, json.JSONDecodeError):
            pid, started_at = -1, ""
        age = time.time() - LOCK_FILE.stat().st_mtime
        if _pid_alive(pid) and age < LOCK_STALE_SECONDS:
            log.info(
                "Lock held by pid=%s since %s (age=%ds). Skipping this run.",
                pid, started_at, int(age),
            )
            return False
        log.warning(
            "Stale lock (pid=%s alive=%s age=%ds) — taking over.",
            pid, _pid_alive(pid), int(age),
        )
    LOCK_FILE.write_text(
        json.dumps({
            "pid": os.getpid(),
            "started_at": datetime.now().isoformat(timespec="seconds"),
        }),
        encoding="utf-8",
    )
    return True


def release_lock() -> None:
    try:
        LOCK_FILE.unlink(missing_ok=True)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Atomic state store
# ---------------------------------------------------------------------------


def load_state() -> dict:
    if not STATE_FILE.exists():
        return {}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        # Recover from a corrupt state file rather than crashing — but keep a
        # copy so the user can inspect it.
        corrupt_copy = STATE_FILE.with_suffix(f".corrupt.{int(time.time())}.json")
        try:
            shutil.copy2(STATE_FILE, corrupt_copy)
            log.warning("processed.json was corrupt; copied to %s and starting empty.", corrupt_copy)
        except OSError:
            log.warning("processed.json was corrupt; starting empty.")
        return {}


def save_state(state: dict) -> None:
    tmp = STATE_FILE.with_suffix(f".tmp.{uuid.uuid4().hex}.json")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, STATE_FILE)


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
        p for p in INBOX_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in ALLOWED_EXTS
    ]
    files.sort(key=lambda p: p.stat().st_mtime)
    return files


# ---------------------------------------------------------------------------
# Text extraction — PDF first, OCR fallback
# ---------------------------------------------------------------------------


def _extract_pdf_text_native(pdf_path: Path) -> str:
    """Use pdfplumber for text-based PDFs. Returns '' if no text found."""
    try:
        import pdfplumber  # type: ignore
    except ImportError as exc:
        raise EnvironmentSkip(
            "pdfplumber not installed. Run: pip install -r requirements.txt"
        ) from exc

    pieces: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pieces.append(text)
            for table in page.extract_tables() or []:
                for row in table:
                    cells = [str(c).strip() for c in row if c is not None]
                    if cells:
                        pieces.append("\t".join(cells))
    return "\n".join(pieces).strip()


def _render_pdf_pages_to_images(pdf_path: Path, out_dir: Path) -> list[Path]:
    try:
        import pypdfium2 as pdfium  # type: ignore
    except ImportError as exc:
        raise EnvironmentSkip(
            "pypdfium2 not installed; cannot render scanned PDFs for OCR fallback. "
            "Run: pip install -r requirements.txt"
        ) from exc

    scale = PDF_OCR_DPI / 72.0
    pages: list[Path] = []
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for idx in range(len(pdf)):
            page = pdf[idx]
            pil_image = page.render(scale=scale).to_pil()
            page_path = out_dir / f"page-{idx + 1:03d}.png"
            pil_image.save(page_path, format="PNG")
            pages.append(page_path)
    finally:
        pdf.close()
    return pages


_OCR_TEXT_RE = re.compile(r"\('((?:[^'\\]|\\.)*)',\s*0?\.\d+\)")


def _run_paddleocr_on_image(image: Path) -> str:
    """Run PaddleOCR in a one-shot container against a single image file.

    The image is staged in a per-call scratch directory so the container only
    ever sees that one file (not the entire inbox).
    """
    with tempfile.TemporaryDirectory(prefix="cowork-ocr-") as scratch:
        scratch_path = Path(scratch)
        staged = scratch_path / image.name
        shutil.copy2(image, staged)

        # Invoke via `python3 -c "import paddleocr; paddleocr.main()"` rather
        # than the `paddleocr` console script — the script in the
        # paddlecloud/paddleocr image looks up its own package metadata via
        # importlib_metadata and that lookup is broken in some builds
        # (PackageNotFoundError: paddleocr).
        cmd = [
            "docker", "run", "--rm",
            "-v", f"{scratch_path.resolve()}:/data",
            "-v", f"{PADDLE_CACHE_DIR.resolve()}:/root/.paddleocr",
            PADDLE_IMAGE,
            "python3", "-c", "import paddleocr; paddleocr.main()",
            "--image_dir", f"/data/{image.name}",
            "--lang", PADDLE_LANG,
            "--use_gpu", "false",
        ]
        log.info("PaddleOCR: %s", " ".join(cmd))
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=SUBPROCESS_TIMEOUT,
            )
        except FileNotFoundError as exc:
            raise EnvironmentSkip("Docker CLI not found on PATH") from exc
        except subprocess.TimeoutExpired as exc:
            raise EnvironmentSkip(
                f"PaddleOCR timed out after {SUBPROCESS_TIMEOUT}s on {image.name}"
            ) from exc

    if result.returncode != 0:
        stderr_tail = (result.stderr or "")[-2000:]
        if "Cannot connect to the Docker daemon" in stderr_tail or "docker daemon" in stderr_tail.lower():
            raise EnvironmentSkip("Docker daemon not reachable")
        raise FileExtractionError(
            f"PaddleOCR failed (exit {result.returncode}): {stderr_tail}"
        )

    matches = _OCR_TEXT_RE.findall(result.stdout)
    if matches:
        return "\n".join(matches)
    # No tuples parsed — return the raw stdout so the LLM still has something.
    return result.stdout.strip()


def _ocr_pdf_pages(pdf_path: Path) -> str:
    with tempfile.TemporaryDirectory(prefix="cowork-pdf-") as scratch:
        pages = _render_pdf_pages_to_images(pdf_path, Path(scratch))
        ocr_pieces = [_run_paddleocr_on_image(p) for p in pages]
    text = "\n".join(s for s in ocr_pieces if s)
    log.info("OCR (scanned PDF, %d page(s)): %d chars", len(ocr_pieces), len(text))
    return text


def extract_text(path: Path) -> str:
    """Return the best available text representation of a bill/statement file.

    For PDFs we run native extraction first (fast + accurate when it works),
    but fall back to OCR if the result is suspiciously short — many stock
    statements embed only their header as real text and rasterise the body,
    which would otherwise leave us with only a few hundred chars to feed the
    LLM.
    """
    ext = path.suffix.lower()
    if ext in PDF_EXTS:
        native = _extract_pdf_text_native(path)
        log.info("PDF native text: %d chars (threshold=%d)", len(native), PDF_TEXT_MIN_CHARS)
        if len(native) >= PDF_TEXT_MIN_CHARS:
            return native
        if native:
            log.warning(
                "PDF native text below threshold (%d < %d); supplementing with OCR.",
                len(native), PDF_TEXT_MIN_CHARS,
            )
        else:
            log.info("PDF had no embedded text; running OCR.")
        ocr_text = _ocr_pdf_pages(path)
        if native and ocr_text:
            # Keep both — native often has clean header values, OCR has the body.
            return native + "\n\n--- OCR ---\n" + ocr_text
        return ocr_text or native
    if ext in IMAGE_EXTS:
        text = _run_paddleocr_on_image(path)
        log.info("OCR (image): %d chars", len(text))
        return text
    raise FileExtractionError(f"Unsupported file extension: {ext}")


# ---------------------------------------------------------------------------
# Ollama extraction with strict JSON schema
# ---------------------------------------------------------------------------

EXTRACTION_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "metadata": {
            "type": "object",
            "properties": {
                "DistributorName": {"type": ["string", "null"]},
                "DivisionName": {"type": ["string", "null"]},
                "ReportFromDate": {"type": ["string", "null"]},
                "ReportToDate": {"type": ["string", "null"]},
                "TotalStockValue": {"type": ["string", "number", "null"]},
                "TotalSalesValue": {"type": ["string", "number", "null"]},
            },
            "required": REPORT_META_COLUMNS,
        },
        "products": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "ProductCode": {"type": ["string", "null"]},
                    "ProductName": {"type": ["string", "null"]},
                    "Packing": {"type": ["string", "null"]},
                    "OpeningStock": {"type": ["string", "number", "null"]},
                    "PurchaseQty": {"type": ["string", "number", "null"]},
                    "SalesQty": {"type": ["string", "number", "null"]},
                    "ClosingStock": {"type": ["string", "number", "null"]},
                    "StockValue": {"type": ["string", "number", "null"]},
                    "SalesValue": {"type": ["string", "number", "null"]},
                    "AdjustmentQty": {"type": ["string", "number", "null"]},
                    "PreviousMonthStock": {"type": ["string", "number", "null"]},
                    "TwoMonthOldStock": {"type": ["string", "number", "null"]},
                    "OldStock120Days": {"type": ["string", "number", "null"]},
                    "ExpiryWithin3Months": {"type": ["string", "number", "null"]},
                    "batches": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "BatchNumber": {"type": ["string", "null"]},
                                "ExpiryDate": {"type": ["string", "null"]},
                                "BatchStock": {"type": ["string", "number", "null"]},
                                "MRP": {"type": ["string", "number", "null"]},
                                "BatchValue": {"type": ["string", "number", "null"]},
                            },
                            "required": BATCH_COLUMNS,
                        },
                    },
                },
                "required": PRODUCT_COLUMNS,
            },
        },
    },
    "required": ["metadata", "products"],
}


SYSTEM_PROMPT = f"""You are an information-extraction engine for distributor stock-statement PDFs.

You will receive the raw text of one stock statement (extracted from a PDF or via OCR). Return STRICT JSON with two keys:

1. "metadata" — object with these fields (use null if not present):
   - DistributorName: company / distributor name on the report header
   - DivisionName: division or business unit, if shown
   - ReportFromDate: start of the reporting period, ISO YYYY-MM-DD when possible else original string
   - ReportToDate: end of the reporting period, same format rules
   - TotalStockValue: report-level total stock value (numeric or string with original formatting)
   - TotalSalesValue: report-level total sales value

2. "products" — array; one element per product/inventory row. Each element has these fields (null when not present):
   - ProductCode
   - ProductName
   - Packing
   - OpeningStock
   - PurchaseQty
   - SalesQty
   - ClosingStock
   - StockValue
   - SalesValue
   - AdjustmentQty
   - PreviousMonthStock
   - TwoMonthOldStock
   - OldStock120Days
   - ExpiryWithin3Months
   - batches: array of batch objects (empty array if no batches for this product)
     Each batch object has: BatchNumber, ExpiryDate, BatchStock, MRP, BatchValue.

Rules:
- Extract EVERY product row visible in the input. Do not summarise or skip rows.
- If batch information is shown for a product, include every batch.
- Dates as ISO YYYY-MM-DD when unambiguous; otherwise keep the original string.
- Numbers: keep as plain numbers without currency symbols (commas as thousands separators are fine inside strings).
- Use null (not "" or "N/A") for fields that are genuinely absent.
- Ignore decorative text, page headers, footers, separators, and page numbers.
- Output JSON only — no markdown fences, no commentary, no trailing prose.
"""


def _compact_text(text: str) -> str:
    """Keep size sensible for the LLM while preserving line order."""
    lines = [ln.rstrip() for ln in text.splitlines()]
    lines = [ln for ln in lines if ln.strip()]
    deduped: list[str] = []
    last = None
    for ln in lines:
        if ln != last:
            deduped.append(ln)
        last = ln
    compact = "\n".join(deduped)
    if len(compact) <= OLLAMA_MAX_TEXT_CHARS:
        return compact
    head_chars = int(OLLAMA_MAX_TEXT_CHARS * 0.7)
    tail_chars = OLLAMA_MAX_TEXT_CHARS - head_chars
    return compact[:head_chars].rstrip() + "\n...\n" + compact[-tail_chars:].lstrip()


def _dump_debug(label: str, content: str) -> Path | None:
    """Write a failure artefact to state/debug/ for after-the-fact inspection."""
    try:
        DEBUG_DIR.mkdir(parents=True, exist_ok=True)
        path = DEBUG_DIR / f"{datetime.now().strftime('%Y%m%d-%H%M%S')}__{label}"
        path.write_text(content, encoding="utf-8")
        return path
    except OSError as exc:
        log.warning("Could not write debug dump %s: %s", label, exc)
        return None


def call_ollama(text: str, source_file: str = "input") -> dict:
    prompt_text = _compact_text(text)
    log.info(
        "Ollama: input text raw=%d chars, compacted=%d chars",
        len(text), len(prompt_text),
    )

    payload = {
        "model": OLLAMA_MODEL,
        "prompt": f"{SYSTEM_PROMPT}\n\nSTOCK_STATEMENT_TEXT:\n<<<\n{prompt_text}\n>>>",
        "stream": False,
        "format": EXTRACTION_SCHEMA,
        "keep_alive": OLLAMA_KEEP_ALIVE,
        "options": {
            "temperature": 0.0,
            "num_ctx": OLLAMA_NUM_CTX,
            "num_predict": OLLAMA_NUM_PREDICT,
        },
    }
    log.info(
        "Ollama: POST /api/generate (model=%s, num_ctx=%d, num_predict=%d, timeout=%ds)",
        OLLAMA_MODEL, OLLAMA_NUM_CTX, OLLAMA_NUM_PREDICT, OLLAMA_TIMEOUT,
    )

    last_err: Exception | None = None
    for attempt in range(1, OLLAMA_RETRY_ATTEMPTS + 1):
        try:
            r = requests.post(
                f"{OLLAMA_URL}/api/generate", json=payload, timeout=OLLAMA_TIMEOUT,
            )
        except (requests.ConnectionError, requests.Timeout) as exc:
            last_err = exc
            log.warning(
                "Ollama transport error (attempt %d/%d): %s",
                attempt, OLLAMA_RETRY_ATTEMPTS, exc,
            )
            if attempt < OLLAMA_RETRY_ATTEMPTS:
                time.sleep(OLLAMA_RETRY_BACKOFF_SEC * attempt)
            continue

        if r.status_code >= 500:
            last_err = RuntimeError(
                f"Ollama HTTP {r.status_code}: {r.text[:300]}"
            )
            log.warning(
                "Ollama server error (attempt %d/%d): %s",
                attempt, OLLAMA_RETRY_ATTEMPTS, last_err,
            )
            if attempt < OLLAMA_RETRY_ATTEMPTS:
                time.sleep(OLLAMA_RETRY_BACKOFF_SEC * attempt)
            continue

        if r.status_code == 404:
            raise EnvironmentSkip(
                f"Ollama model '{OLLAMA_MODEL}' not pulled. "
                f"Run: docker exec cowork-ollama ollama pull {OLLAMA_MODEL}"
            )

        r.raise_for_status()
        body = r.json()
        raw = (body.get("response") or "").strip()
        done_reason = body.get("done_reason") or ""
        eval_count = body.get("eval_count")
        log.info(
            "Ollama: response len=%d chars, eval_count=%s, done_reason=%s",
            len(raw), eval_count, done_reason,
        )
        try:
            return json.loads(raw)
        except json.JSONDecodeError as exc:
            m = re.search(r"\{.*\}", raw, re.DOTALL)
            if not m:
                dump = _dump_debug(f"{source_file}.no-json.txt", raw)
                raise FileExtractionError(
                    f"Ollama response was not JSON (len={len(raw)}). "
                    f"Full output dumped to: {dump}"
                ) from exc
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError as exc2:
                dump = _dump_debug(f"{source_file}.partial-json.txt", raw)
                hint = ""
                if done_reason == "length" or (
                    eval_count is not None and eval_count >= OLLAMA_NUM_PREDICT - 16
                ):
                    hint = (
                        " — output hit num_predict ceiling; raise OLLAMA_NUM_PREDICT "
                        "or split the source into smaller pieces."
                    )
                raise FileExtractionError(
                    f"Ollama response could not be parsed as JSON "
                    f"(len={len(raw)}, eval_count={eval_count}, done_reason={done_reason}){hint}. "
                    f"Full output dumped to: {dump}"
                ) from exc2

    raise EnvironmentSkip(
        f"Ollama unreachable after {OLLAMA_RETRY_ATTEMPTS} attempt(s): {last_err}"
    )


# ---------------------------------------------------------------------------
# Schema validation + row flattening
# ---------------------------------------------------------------------------


@dataclass
class ExtractedDoc:
    metadata: dict
    products: list[dict]


def _coerce_scalar(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return value
    return str(value).strip()


def validate_and_normalize(raw: dict) -> ExtractedDoc:
    if not isinstance(raw, dict):
        raise FileExtractionError("LLM output is not a JSON object")

    metadata = raw.get("metadata") or {}
    if not isinstance(metadata, dict):
        raise FileExtractionError("metadata is not an object")

    products = raw.get("products") or []
    if not isinstance(products, list):
        raise FileExtractionError("products is not an array")

    norm_meta = {col: _coerce_scalar(metadata.get(col)) for col in REPORT_META_COLUMNS}

    norm_products: list[dict] = []
    for idx, product in enumerate(products):
        if not isinstance(product, dict):
            raise FileExtractionError(f"products[{idx}] is not an object")
        norm = {col: _coerce_scalar(product.get(col)) for col in PRODUCT_COLUMNS}
        batches_raw = product.get("batches") or []
        if not isinstance(batches_raw, list):
            raise FileExtractionError(f"products[{idx}].batches is not an array")
        norm_batches: list[dict] = []
        for bidx, batch in enumerate(batches_raw):
            if not isinstance(batch, dict):
                raise FileExtractionError(
                    f"products[{idx}].batches[{bidx}] is not an object"
                )
            norm_batches.append(
                {col: _coerce_scalar(batch.get(col)) for col in BATCH_COLUMNS}
            )
        norm["batches"] = norm_batches
        norm_products.append(norm)

    if not norm_products:
        raise FileExtractionError("LLM returned zero products")

    return ExtractedDoc(metadata=norm_meta, products=norm_products)


def flatten_rows(doc: ExtractedDoc, source_file: str, processed_at: str) -> list[list[Any]]:
    rows: list[list[Any]] = []
    blank_batch = {col: "" for col in BATCH_COLUMNS}
    for product in doc.products:
        batches = product.get("batches") or []
        if not batches:
            batches = [blank_batch]
        for batch in batches:
            row = [processed_at, source_file]
            row.extend(doc.metadata.get(col, "") for col in REPORT_META_COLUMNS)
            row.extend(product.get(col, "") for col in PRODUCT_COLUMNS)
            row.extend(batch.get(col, "") for col in BATCH_COLUMNS)
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Atomic Excel append
# ---------------------------------------------------------------------------


def ensure_output_writable() -> None:
    if not OUTPUT_XLSX.exists():
        return
    try:
        with OUTPUT_XLSX.open("a+b"):
            pass
    except PermissionError as exc:
        raise EnvironmentSkip(
            f"{OUTPUT_XLSX.name} is open in Excel — close it and re-run."
        ) from exc


def _load_or_create_workbook() -> tuple[Workbook, Any]:
    if OUTPUT_XLSX.exists():
        wb = load_workbook(OUTPUT_XLSX)
        ws = wb.active
        # Heal a workbook missing the header row.
        if ws.max_row == 0 or [c.value for c in ws[1]] != EXCEL_COLUMNS:
            if ws.max_row == 0:
                ws.append(EXCEL_COLUMNS)
        return wb, ws
    wb = Workbook()
    ws = wb.active
    ws.title = "inventory"
    ws.append(EXCEL_COLUMNS)
    return wb, ws


def append_rows_atomic(rows: list[list[Any]]) -> int:
    """Append rows or fail without touching the existing workbook.

    Strategy: load (or create) → append in memory → save to .new → fsync →
    backup current to .bak → os.replace .new over the live file → drop .bak.
    If anything throws before the final replace, the live file is untouched.
    """
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

    wb, ws = _load_or_create_workbook()
    for row in rows:
        ws.append(row)

    tmp_path = OUTPUT_XLSX.with_suffix(f".new.{uuid.uuid4().hex}.xlsx")
    bak_path = OUTPUT_XLSX.with_suffix(".bak.xlsx")

    buf = io.BytesIO()
    wb.save(buf)
    data = buf.getvalue()

    with tmp_path.open("wb") as f:
        f.write(data)
        f.flush()
        os.fsync(f.fileno())

    backed_up = False
    if OUTPUT_XLSX.exists():
        shutil.copy2(OUTPUT_XLSX, bak_path)
        backed_up = True

    try:
        os.replace(tmp_path, OUTPUT_XLSX)
    except PermissionError as exc:
        # Workbook got opened in Excel between our writable check and the swap.
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise EnvironmentSkip(
            f"{OUTPUT_XLSX.name} became locked during write; will retry next run."
        ) from exc

    if backed_up:
        try:
            bak_path.unlink(missing_ok=True)
        except OSError:
            pass
    return len(rows)


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
# Notifications
# ---------------------------------------------------------------------------


def notify(text: str) -> None:
    if not NOTIFY_WEBHOOK_URL:
        return
    try:
        requests.post(NOTIFY_WEBHOOK_URL, json={"text": text}, timeout=10)
    except requests.RequestException as exc:
        log.warning("Notification webhook failed: %s", exc)


# ---------------------------------------------------------------------------
# Per-file pipeline
# ---------------------------------------------------------------------------


@dataclass
class RunStats:
    succeeded: int = 0
    file_failed: int = 0
    env_skipped: int = 0
    moved_to_failed: int = 0


def process_one(path: Path, state: dict, stats: RunStats) -> None:
    log.info("---- processing %s ----", path.name)
    digest = sha256_of(path)
    record = state.get(digest, {}) or {}

    if record.get("status") == "success":
        log.info(
            "Already processed (hash=%s, on %s). Moving duplicate to processed/.",
            digest[:10], record.get("processed_at"),
        )
        try:
            move_to(path, PROCESSED_DIR)
        except OSError as exc:
            log.warning("Could not move duplicate: %s", exc)
        return

    attempts_so_far = int(record.get("attempts", 0))

    try:
        ensure_output_writable()
        text = extract_text(path)
        if not text.strip():
            raise FileExtractionError("Extracted text was empty")

        raw = call_ollama(text, source_file=path.name)
        doc = validate_and_normalize(raw)

        processed_at = datetime.now().isoformat(timespec="seconds")
        rows = flatten_rows(doc, source_file=path.name, processed_at=processed_at)
        if not rows:
            raise FileExtractionError("No rows produced after flattening")

        ensure_output_writable()
        n_written = append_rows_atomic(rows)
        dest = move_to(path, PROCESSED_DIR)

        state[digest] = {
            "source_file": path.name,
            "processed_at": processed_at,
            "status": "success",
            "rows_added": n_written,
            "products": len(doc.products),
            "attempts": attempts_so_far + 1,
            "moved_to": str(dest),
        }
        save_state(state)
        stats.succeeded += 1
        log.info(
            "OK: %d row(s) from %d product(s) appended to %s",
            n_written, len(doc.products), OUTPUT_XLSX,
        )

    except EnvironmentSkip as exc:
        # Worker-level problem; do NOT burn an attempt.
        stats.env_skipped += 1
        log.warning("SKIP (environment): %s", exc)

    except FileExtractionError as exc:
        new_attempts = attempts_so_far + 1
        log.error(
            "FAIL (%d/%d): %s — %s",
            new_attempts, RETRY_LIMIT, path.name, exc,
        )
        stats.file_failed += 1

        if new_attempts >= RETRY_LIMIT:
            try:
                dest = move_to(path, FAILED_DIR)
                moved_to = str(dest)
            except OSError as move_exc:
                log.warning("Could not move to failed/: %s", move_exc)
                moved_to = str(path)
            state[digest] = {
                "source_file": path.name,
                "last_attempt_at": datetime.now().isoformat(timespec="seconds"),
                "status": "failed",
                "attempts": new_attempts,
                "error": str(exc),
                "moved_to": moved_to,
            }
            save_state(state)
            stats.moved_to_failed += 1
            notify(
                f"Cowork: {path.name} moved to failed/ after {new_attempts} attempts. "
                f"Last error: {exc}"
            )
        else:
            state[digest] = {
                "source_file": path.name,
                "last_attempt_at": datetime.now().isoformat(timespec="seconds"),
                "status": "retrying",
                "attempts": new_attempts,
                "error": str(exc),
            }
            save_state(state)

    except Exception as exc:  # noqa: BLE001
        # Unexpected programming error — treat as file-specific so we don't
        # silently retry forever, but log loud.
        new_attempts = attempts_so_far + 1
        log.exception("UNEXPECTED FAIL (%d/%d) %s", new_attempts, RETRY_LIMIT, path.name)
        stats.file_failed += 1
        if new_attempts >= RETRY_LIMIT:
            try:
                dest = move_to(path, FAILED_DIR)
                moved_to = str(dest)
            except OSError:
                moved_to = str(path)
            state[digest] = {
                "source_file": path.name,
                "last_attempt_at": datetime.now().isoformat(timespec="seconds"),
                "status": "failed",
                "attempts": new_attempts,
                "error": f"{type(exc).__name__}: {exc}",
                "moved_to": moved_to,
            }
            save_state(state)
            stats.moved_to_failed += 1
            notify(
                f"Cowork: {path.name} moved to failed/ after {new_attempts} attempts. "
                f"Unexpected error: {type(exc).__name__}: {exc}"
            )
        else:
            state[digest] = {
                "source_file": path.name,
                "last_attempt_at": datetime.now().isoformat(timespec="seconds"),
                "status": "retrying",
                "attempts": new_attempts,
                "error": f"{type(exc).__name__}: {exc}",
            }
            save_state(state)


# ---------------------------------------------------------------------------
# Status CLI
# ---------------------------------------------------------------------------


def print_status() -> int:
    state = load_state()
    if not state:
        print("No files processed yet.")
        return 0
    by_status: dict[str, int] = {}
    retrying: list[tuple[int, str, str]] = []
    failed: list[tuple[str, str]] = []
    succeeded_today = 0
    today = datetime.now().date().isoformat()

    for record in state.values():
        status = record.get("status", "unknown")
        by_status[status] = by_status.get(status, 0) + 1
        if status == "success" and (record.get("processed_at") or "").startswith(today):
            succeeded_today += 1
        elif status == "retrying":
            retrying.append((
                int(record.get("attempts", 0)),
                record.get("source_file", "?"),
                record.get("error", ""),
            ))
        elif status == "failed":
            failed.append((record.get("source_file", "?"), record.get("error", "")))

    print(f"Total tracked files: {sum(by_status.values())}")
    for status, count in sorted(by_status.items()):
        print(f"  {status}: {count}")
    print(f"  succeeded today: {succeeded_today}")

    if retrying:
        print("\nIn retry queue:")
        for attempts, name, err in sorted(retrying, reverse=True):
            print(f"  [{attempts}/{RETRY_LIMIT}] {name} — {err[:120]}")
    if failed:
        print("\nMoved to failed/ (latest first):")
        for name, err in failed[-10:]:
            print(f"  {name} — {err[:120]}")
    return 0


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


@contextmanager
def _lock_scope():
    if not acquire_lock():
        yield False
        return
    try:
        yield True
    finally:
        release_lock()


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if args and args[0] in {"--status", "status"}:
        return print_status()

    run_started = time.monotonic()
    log.info("Cowork worker run starting. INBOX=%s MODEL=%s", INBOX_DIR, OLLAMA_MODEL)

    with _lock_scope() as got_lock:
        if not got_lock:
            return 0
        stats = RunStats()
        files = discover_inbox()
        log.info("Inbox: %d candidate file(s)", len(files))
        if not files:
            log.info(
                "RUN_SUMMARY processed=0 succeeded=0 file_failed=0 env_skipped=0 "
                "moved_to_failed=0 duration=%.1fs",
                time.monotonic() - run_started,
            )
            return 0

        if BATCH_LIMIT > 0:
            files = files[:BATCH_LIMIT]

        state = load_state()
        for f in files:
            process_one(f, state, stats)

        log.info(
            "RUN_SUMMARY processed=%d succeeded=%d file_failed=%d env_skipped=%d "
            "moved_to_failed=%d duration=%.1fs",
            len(files),
            stats.succeeded,
            stats.file_failed,
            stats.env_skipped,
            stats.moved_to_failed,
            time.monotonic() - run_started,
        )
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
