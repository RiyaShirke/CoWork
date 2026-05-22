# One-time bootstrap for the Cowork worker.
#   powershell -ExecutionPolicy Bypass -File <project>\scripts\setup_docker.ps1
#
# What it does:
#   1. Installs the Python dependencies from requirements.txt.
#   2. Pulls the PaddleOCR image (used per-file via `docker run --rm`).
#   3. Brings up the Ollama container in the background via docker-compose.
#   4. Pulls the configured Ollama model inside that container.
#
# Override the model by setting $env:OLLAMA_MODEL before running, e.g.:
#   $env:OLLAMA_MODEL = 'llama3.1:8b'; .\setup_docker.ps1

$ErrorActionPreference = 'Stop'

# Resolve project root from this script's location so the script works from
# any drive / install path without editing.
$ScriptDir       = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir      = Split-Path -Parent $ScriptDir
$ComposeFile     = Join-Path $ProjectDir 'docker-compose.yml'
$RequirementsTxt = Join-Path $ProjectDir 'requirements.txt'

if (-not (Test-Path $ComposeFile)) {
    throw "docker-compose.yml not found at $ComposeFile"
}
if (-not (Test-Path $RequirementsTxt)) {
    throw "requirements.txt not found at $RequirementsTxt"
}

$Model = if ($env:OLLAMA_MODEL) { $env:OLLAMA_MODEL } else { 'llama3.1:8b' }
$PaddleImage = if ($env:PADDLE_IMAGE) { $env:PADDLE_IMAGE } else { 'paddlecloud/paddleocr:2.6-cpu-latest' }

# Prefer `py -3` if available (the Windows launcher); fall back to `python`.
$PythonCmd = $null
foreach ($candidate in @('py', 'python')) {
    $resolved = Get-Command $candidate -ErrorAction SilentlyContinue
    if ($resolved) { $PythonCmd = $resolved.Source; break }
}
if (-not $PythonCmd) {
    throw "No Python interpreter found on PATH (looked for 'py' and 'python')."
}

Write-Host "==> Installing Python dependencies from $RequirementsTxt"
Write-Host "    using: $PythonCmd -m pip install -r requirements.txt"
& $PythonCmd -m pip install --disable-pip-version-check -r $RequirementsTxt
if ($LASTEXITCODE -ne 0) { throw "pip install failed (exit $LASTEXITCODE)" }

Write-Host ""
Write-Host "==> Pulling PaddleOCR image: $PaddleImage"
docker pull $PaddleImage
if ($LASTEXITCODE -ne 0) { throw "docker pull failed for $PaddleImage" }

Write-Host ""
Write-Host "==> Bringing up Ollama (docker compose)"
docker compose -f $ComposeFile up -d
if ($LASTEXITCODE -ne 0) { throw "docker compose up failed" }

Write-Host ""
Write-Host "==> Waiting for Ollama to accept connections on :11434"
$ready = $false
for ($i = 0; $i -lt 30; $i++) {
    try {
        $r = Invoke-WebRequest -Uri 'http://localhost:11434/api/tags' -UseBasicParsing -TimeoutSec 2
        if ($r.StatusCode -eq 200) { Write-Host "    ready."; $ready = $true; break }
    } catch { Start-Sleep -Seconds 2 }
}
if (-not $ready) { throw "Ollama did not become ready within 60s" }

Write-Host ""
Write-Host "==> Pulling Ollama model: $Model (this may take a few minutes)"
docker exec cowork-ollama ollama pull $Model
if ($LASTEXITCODE -ne 0) { throw "ollama pull failed for $Model" }

Write-Host ""
Write-Host "Bootstrap complete."
Write-Host "  Python          : $PythonCmd"
Write-Host "  PaddleOCR image : $PaddleImage"
Write-Host "  Ollama model    : $Model"
Write-Host "  Ollama endpoint : http://localhost:11434"
Write-Host "  Project dir     : $ProjectDir"
Write-Host ""
Write-Host "Next: drop a stock-statement PDF into $(Join-Path $ProjectDir 'inbox') and run:"
Write-Host "    $PythonCmd $(Join-Path $ProjectDir 'worker.py')"
