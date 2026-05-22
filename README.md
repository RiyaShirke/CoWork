# Cowork — local stock-statement ingestion agent

A small, local agent that watches a folder for **stock-statement PDFs**
(distributor inventory reports), extracts every product/batch row with a local
LLM, and appends them to a single Excel workbook. Runs every 15 minutes via
Windows Task Scheduler. No database, no cloud.

```
D:\Cowork\
├── worker.py                # the pipeline (one file)
├── docker-compose.yml       # Ollama service
├── .env / .env.example      # configuration
├── requirements.txt
├── inbox\                   # drop stock-statement PDFs / images here
├── processed\               # successful files move here (timestamp-prefixed)
├── failed\                  # files that exhausted RETRY_LIMIT attempts
├── output\bills.xlsx        # the single rolling Excel output (atomic writes)
├── state\processed.json     # SHA256 -> {status, attempts, ...}
├── state\worker.lock        # active-run lock (PID-aware)
├── logs\cowork.log          # rolling log (5 MB × 5 files)
└── scripts\
    ├── run_worker.cmd       # entry point used by Task Scheduler
    ├── register_task.ps1    # register the scheduled task (run as admin)
    ├── unregister_task.ps1  # remove it
    └── setup_docker.ps1     # one-time: pull images, start Ollama, pull model
```

## Atomicity guarantees

Each file's pipeline is **all-or-nothing**:

1. The Excel workbook is updated *and* the state file is updated *and* the
   source file is moved to `processed\` — or none of that happens.
2. `output\bills.xlsx` is written to a temp file, fsynced, then `os.replace`d
   over the live file. The original is also copied to `.bak` immediately
   beforehand. A crash leaves the live workbook either untouched or fully
   updated, never half-written.
3. `state\processed.json` is written via tmp-file + fsync + `os.replace`.
4. The PID-aware lock prevents overlapping runs but never strands a crashed
   worker for the full 1-hour stale window (the PID is checked for liveness).

## Retry policy

- Failures are split into two classes:
  - **File-specific failures** (corrupt PDF, OCR returned nothing, LLM output
    failed schema validation): increment `attempts` for that file's hash.
    After `RETRY_LIMIT` (default 3) attempts the file is moved to `failed\`.
  - **Environment failures** (Docker daemon down, Ollama unreachable,
    `bills.xlsx` open in Excel, OCR subprocess timed out): the run is
    skipped — the file stays in `inbox\` and **no attempt is burned**.
- Retries happen on the next scheduled run; no per-file backoff timer.
- When a file is moved to `failed\`, an optional webhook (`NOTIFY_WEBHOOK_URL`)
  is pinged so a human can investigate.

## Extracted columns

The Excel file has 27 columns. One row per `(product, batch)` pair (one row
per product when the statement has no batch info, with batch cols blank).
Report-level metadata is repeated on every row for trivial pivoting.

| Group           | Columns                                                                                                                                                                                            |
|-----------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Provenance      | `processed_at`, `source_file`                                                                                                                                                                      |
| Report metadata | `DistributorName`, `DivisionName`, `ReportFromDate`, `ReportToDate`, `TotalStockValue`, `TotalSalesValue`                                                                                          |
| Product         | `ProductCode`, `ProductName`, `Packing`, `OpeningStock`, `PurchaseQty`, `SalesQty`, `ClosingStock`, `StockValue`, `SalesValue`, `AdjustmentQty`, `PreviousMonthStock`, `TwoMonthOldStock`, `OldStock120Days`, `ExpiryWithin3Months` |
| Batch (if any)  | `BatchNumber`, `ExpiryDate`, `BatchStock`, `MRP`, `BatchValue`                                                                                                                                     |

Missing values are written as empty strings. The LLM is instructed to extract
every product row visible in the input — no summarising, no skipping.

## How a run works

1. Acquire `state\worker.lock` (skip if another live PID owns it).
2. List `inbox\`, sorted oldest-first; take the first `BATCH_LIMIT` files.
3. For each file:
   - SHA256 the content. If already `status=success`, move it to `processed\`
     and skip — same content never produces duplicate rows.
   - **Extract text**: for PDFs, try `pdfplumber` first (fast + accurate on
     text-based PDFs). If the PDF has no embedded text, render each page to
     PNG via `pypdfium2` and OCR with PaddleOCR (Docker). Image files go
     straight to PaddleOCR. The OCR container only sees the single file via a
     scratch directory — not the whole inbox.
   - **Call Ollama** at `/api/generate` with `format=<JSON schema>` so the
     model returns strict JSON. Retries on transport/5xx errors with backoff.
   - **Validate** the response against the schema; flatten products & batches
     into rows.
   - **Atomically append** rows to `output\bills.xlsx`, then update state,
     then move the source file to `processed\`. If any step fails, none of
     them happen — see *Atomicity guarantees* above.
4. Release the lock and log a `RUN_SUMMARY` line.

## Setup (one-time)

1. **Bootstrap** (installs Python deps + Docker images + Ollama + model in one go):
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\setup_docker.ps1
   ```
   This `pip install`s `requirements.txt` (including **Kreuzberg** for
   table-aware PDF extraction), pulls `paddlecloud/paddleocr:2.6-cpu-latest`,
   starts the `cowork-ollama` container on `localhost:11434`, and pulls
   `qwen2.5:7b`. Override the model with `$env:OLLAMA_MODEL = 'llama3.1:8b'`
   etc. before running. Anything 3b or smaller will drop fields on a
   27-column schema with nested batches — don't use it.

   **GPU is auto-detected.** If `nvidia-smi` works on the host the script
   also applies `docker-compose.gpu.yml` so Ollama loads weights into VRAM
   (much faster, much less system RAM pressure). It then verifies the GPU
   is actually visible inside the container and prints
   `Inference mode: GPU (<card name>)` at the end. Force CPU even on a GPU
   host with `$env:COWORK_FORCE_CPU = '1'` before running.

2. **Register the scheduled task** (every 15 min):
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\register_task.ps1
   ```

## Daily use

Drop stock statements into `inbox\`. Within 15 minutes (or
`Start-ScheduledTask -TaskName CoworkWorker` to fire immediately) they'll be
processed and appended to `output\bills.xlsx`.

## Status check

```powershell
python worker.py --status
```

Prints totals per status (`success`, `retrying`, `failed`), today's success
count, the current retry queue with attempt counters, and the most recent
files that exhausted `RETRY_LIMIT`.

## Manual run (no scheduler)

```powershell
python worker.py
```

## Configuration

All knobs live in `.env` (see `.env.example`). The important ones:

| Variable                  | Default                                | Meaning                                              |
|---------------------------|----------------------------------------|------------------------------------------------------|
| `INBOX_DIR`               | `D:\Cowork\inbox`                      | Folder watched for new statements                    |
| `OUTPUT_XLSX`             | `D:\Cowork\output\bills.xlsx`          | Excel file rows are appended to                      |
| `BATCH_LIMIT`             | `1`                                    | Max files per run (`0` = all queued)                 |
| `RETRY_LIMIT`             | `3`                                    | File-specific failures before move to `failed\`      |
| `SUBPROCESS_TIMEOUT`      | `600`                                  | Hard timeout (sec) for OCR Docker subprocess         |
| `OLLAMA_MODEL`            | `llama3.1:8b`                          | Local LLM. Must handle 14+ fields reliably.          |
| `OLLAMA_NUM_CTX`          | `8192`                                 | Context window passed to Ollama                      |
| `OLLAMA_MAX_TEXT_CHARS`   | `24000`                                | Source text cap before sending to LLM                |
| `OLLAMA_RETRY_ATTEMPTS`   | `3`                                    | Transport retries on connection/timeout/5xx          |
| `PDF_OCR_DPI`             | `200`                                  | DPI when rendering scanned PDFs for OCR              |
| `NOTIFY_WEBHOOK_URL`      | *(empty)*                              | Optional Slack/Discord/generic webhook on `failed\`  |

## Idempotency

`state\processed.json` maps `sha256 → { status, attempts, processed_at, ... }`.
If the same content lands in the inbox again it's moved straight to
`processed\` without re-running OCR / LLM. Files in `retrying` get re-attempted
every scheduled run until they succeed or hit `RETRY_LIMIT`.

To force a re-attempt of a file already in `failed\`: delete its hash entry
from `processed.json`, move the file back into `inbox\`.

## Troubleshooting

- **No Ollama**: `docker ps` should show `cowork-ollama`. If not, rerun
  `scripts\setup_docker.ps1`. The worker logs `SKIP (environment): Ollama
  unreachable…` and never burns an attempt for this.
- **Model not pulled**: the worker logs the exact `docker exec cowork-ollama
  ollama pull <model>` to run.
- **`docker: command not found` in scheduler**: ensure Docker Desktop starts
  at login, and the task's principal is the same user as Docker Desktop.
- **Workbook open in Excel**: the run skips (no attempt burned) and logs a
  `SKIP (environment)` message. Close Excel; next tick picks it up.
- **PDF produced no text**: the worker falls back to OCR. If OCR also returns
  empty, the attempt is burned and the file moves to `failed\` after 3 tries.
- **Looks stuck**: `python worker.py --status` shows what's in the retry queue
  and the latest error per file.
