# Cowork — local AI bill-ingestion agent

A small, local agent that watches a folder for bill images / PDFs, OCRs them
with **PaddleOCR** (Docker), extracts structured rows with a **local Ollama**
LLM (Docker), and appends them to a single Excel workbook. Runs every
15 minutes via **Windows Task Scheduler**. No database, no cloud.

```
D:\Cowork\
├── worker.py              # the pipeline (one file)
├── docker-compose.yml     # Ollama service
├── .env / .env.example    # configuration
├── requirements.txt
├── inbox\                 # drop bill images / PDFs here
├── processed\             # successful files move here (timestamp-prefixed)
├── failed\                # files whose pipeline errored move here
├── output\bills.xlsx      # the single rolling Excel output
├── state\processed.json   # SHA256 -> result (idempotency)
├── state\worker.lock      # active-run lock
├── logs\cowork.log        # rolling log
└── scripts\
    ├── run_worker.cmd       # entry point used by Task Scheduler
    ├── register_task.ps1    # register the scheduled task (run as admin)
    ├── unregister_task.ps1  # remove it
    └── setup_docker.ps1     # one-time: pull images, start Ollama, pull model
```

## How a run works

1. Acquire `state\worker.lock` (skip if another run is active).
2. List `inbox\` for allowed extensions, sorted oldest-first.
3. Take the first `BATCH_LIMIT` files (default 1).
4. For each file:
   - SHA256 the content. If already in `processed.json` as `success`, move to
     `processed\` and skip — guarantees the same image is never processed twice.
   - Run PaddleOCR in a one-shot container:
     `docker run --rm -v <inbox>:/data paddlecloud/paddleocr ... /data/<file>`
   - Send the OCR text to Ollama (`/api/generate`, `format=json`) asking for a
     `{"rows": [...]}` payload with vendor / invoice_no / date / line_item /
     quantity / unit_price / amount / currency / notes.
   - Append each row to `output\bills.xlsx` (file is created on first run).
   - Move the source file to `processed\` (or `failed\` on any error) and
     record the result in `processed.json`.
5. Release the lock.

## Setup (one-time)

1. **Install Python deps** (already done if you ran the bootstrap):
   ```powershell
   pip install -r D:\Cowork\requirements.txt
   ```

2. **Bring up Docker** (PaddleOCR image + Ollama service + model):
   ```powershell
   powershell -ExecutionPolicy Bypass -File D:\Cowork\scripts\setup_docker.ps1
   ```
   This pulls `paddlecloud/paddleocr:2.6-cpu-latest`, starts the
  `cowork-ollama` container on `localhost:11434`, and pulls `llama3.2:1b`.
  Override the model with `$env:OLLAMA_MODEL = 'llama3.2:3b'` before running.

3. **Register the scheduled task** (every 15 min):
   ```powershell
   powershell -ExecutionPolicy Bypass -File D:\Cowork\scripts\register_task.ps1
   ```
   Run as the same user who will own the inbox.

## Daily use

Drop bills into `D:\Cowork\inbox\`. Within 15 minutes (or run
`Start-ScheduledTask -TaskName CoworkWorker` to fire immediately) they will be
processed and the rows appended to `D:\Cowork\output\bills.xlsx`.

## Manual run (no scheduler)

```powershell
cd D:\Cowork
python worker.py
```

## Configuration

All knobs live in `.env` (see `.env.example`). The important ones:

| Variable        | Default                                       | Meaning                                     |
|-----------------|-----------------------------------------------|---------------------------------------------|
| `INBOX_DIR`     | `D:\Cowork\inbox`                             | Folder watched for new bills                |
| `OUTPUT_XLSX`   | `D:\Cowork\output\bills.xlsx`                 | Single Excel file rows are appended to      |
| `BATCH_LIMIT`   | `1`                                           | Max files per run (`0` = all available)     |
| `ALLOWED_EXTS`  | `.png,.jpg,.jpeg,.bmp,.tif,.tiff,.pdf`        | File extensions considered                  |
| `PADDLE_IMAGE`  | `paddlecloud/paddleocr:2.6-cpu-latest`        | OCR Docker image                            |
| `OLLAMA_MODEL`  | `llama3.2:1b`                                 | Local LLM used for faster extraction        |
| `OLLAMA_URL`    | `http://localhost:11434`                      | Ollama HTTP endpoint                        |

Override per-machine by editing `D:\Cowork\.env`, or per-run with normal
environment variables — `$env:INBOX_DIR = 'E:\incoming'` etc.

## Idempotency

`state\processed.json` maps `sha256 → { status, processed_at, rows_added }`.
If the same image (same bytes) lands in the inbox again, it is moved straight
to `processed\` without re-running OCR / LLM. Failed files are still recorded
(status `failed`) so you can inspect them in `failed\` and decide whether to
retry — to retry, delete the file's hash entry from `processed.json` and drop
the file back into `inbox\`.

## Troubleshooting

- **No Ollama**: `docker ps` should show `cowork-ollama`. If not, rerun
  `scripts\setup_docker.ps1`.
- **`docker: command not found` in scheduler**: ensure Docker Desktop is set to
  start at login, and the task's principal is the same user as Docker Desktop.
- **PaddleOCR returns nothing**: check `logs\cowork.log` — the raw `docker run`
  stdout is in there. Adjust `PADDLE_LANG` if your bills aren't English.
- **Model output isn't JSON**: the worker tries `format=json` + a regex
  fallback; if it still fails, try a stronger model (`llama3.1:8b`).
- **Two runs collide**: `state\worker.lock` blocks overlap. Stale locks
  (>1 hour old) are auto-taken-over.
