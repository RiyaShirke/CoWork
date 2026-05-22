# One-time bootstrap for the Cowork worker.
#   powershell -ExecutionPolicy Bypass -File <project>\scripts\setup_docker.ps1
#
# What it does:
#   1. Installs the Python dependencies from requirements.txt.
#   2. Pulls the PaddleOCR image (used per-file via `docker run --rm`).
#   3. Detects an NVIDIA GPU on the host (via nvidia-smi). If present, the
#      Ollama container is brought up with GPU access via the
#      docker-compose.gpu.yml overlay; otherwise CPU-only.
#   4. Pulls the configured Ollama model inside that container.
#
# Override the model by setting $env:OLLAMA_MODEL before running, e.g.:
#   $env:OLLAMA_MODEL = 'llama3.1:8b'; .\setup_docker.ps1
#
# Force CPU even on a GPU host:
#   $env:COWORK_FORCE_CPU = '1'; .\setup_docker.ps1

$ErrorActionPreference = 'Stop'

# Resolve project root from this script's location so the script works from
# any drive / install path without editing.
$ScriptDir       = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir      = Split-Path -Parent $ScriptDir
$ComposeFile     = Join-Path $ProjectDir 'docker-compose.yml'
$ComposeGpuFile  = Join-Path $ProjectDir 'docker-compose.gpu.yml'
$RequirementsTxt = Join-Path $ProjectDir 'requirements.txt'

if (-not (Test-Path $ComposeFile)) {
    throw "docker-compose.yml not found at $ComposeFile"
}
if (-not (Test-Path $RequirementsTxt)) {
    throw "requirements.txt not found at $RequirementsTxt"
}

$Model = if ($env:OLLAMA_MODEL) { $env:OLLAMA_MODEL } else { 'qwen2.5:7b' }
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
Write-Host "==> Pulling Kreuzberg image: ghcr.io/kreuzberg-dev/kreuzberg:latest"
docker pull ghcr.io/kreuzberg-dev/kreuzberg:latest
if ($LASTEXITCODE -ne 0) { throw "docker pull failed for Kreuzberg" }

# --- GPU detection -----------------------------------------------------------
$UseGpu = $false
$GpuName = ''
if ($env:COWORK_FORCE_CPU -eq '1') {
    Write-Host ""
    Write-Host "==> COWORK_FORCE_CPU=1 set; skipping GPU detection."
} else {
    $nvidia = Get-Command 'nvidia-smi' -ErrorAction SilentlyContinue
    if ($nvidia) {
        try {
            $GpuName = (& $nvidia.Source --query-gpu=name --format=csv,noheader 2>$null | Select-Object -First 1)
            if ($LASTEXITCODE -eq 0 -and $GpuName) {
                $UseGpu = $true
                Write-Host ""
                Write-Host "==> NVIDIA GPU detected: $GpuName"
                if (-not (Test-Path $ComposeGpuFile)) {
                    Write-Host "    WARNING: $ComposeGpuFile missing; falling back to CPU."
                    $UseGpu = $false
                }
            }
        } catch {
            Write-Host "==> nvidia-smi exists but failed to query; falling back to CPU."
            $UseGpu = $false
        }
    } else {
        Write-Host ""
        Write-Host "==> No nvidia-smi on PATH; running Ollama CPU-only."
    }
}

# --- Bring up services -------------------------------------------------------
Write-Host ""
if ($UseGpu) {
    Write-Host "==> Bringing up services with Ollama GPU passthrough (base + .gpu overlay)"
    docker compose -f $ComposeFile -f $ComposeGpuFile up -d
} else {
    Write-Host "==> Bringing up services (Ollama on CPU)"
    docker compose -f $ComposeFile up -d
}
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
Write-Host "==> Waiting for Kreuzberg to accept connections on :8000"
$kzReady = $false
for ($i = 0; $i -lt 30; $i++) {
    try {
        $r = Invoke-WebRequest -Uri 'http://localhost:8000/health' -UseBasicParsing -TimeoutSec 2
        if ($r.StatusCode -eq 200) { Write-Host "    ready."; $kzReady = $true; break }
    } catch { Start-Sleep -Seconds 2 }
}
if (-not $kzReady) {
    Write-Host "    WARNING: Kreuzberg did not become ready within 60s — worker will fall back to pdfplumber."
}

# Verify GPU actually passed through into the container (otherwise the model
# loads on CPU silently, which is the worst of both worlds).
$GpuInContainer = $false
if ($UseGpu) {
    Write-Host ""
    Write-Host "==> Verifying GPU is visible inside the Ollama container"
    $smi = docker exec cowork-ollama nvidia-smi --query-gpu=name --format=csv,noheader 2>$null
    if ($LASTEXITCODE -eq 0 -and $smi) {
        Write-Host "    GPU OK inside container: $smi"
        $GpuInContainer = $true
    } else {
        Write-Host "    Host has a GPU but the container can't see it."
        Write-Host "    Likely cause: Docker Desktop is missing the NVIDIA Container Toolkit / WSL2 GPU support."
        Write-Host "    The model will load on CPU. To fix later, ensure Docker Desktop is up to date and the NVIDIA Windows driver is recent, then re-run setup."
    }
}

Write-Host ""
Write-Host "==> Pulling Ollama model: $Model (this may take a few minutes)"
docker exec cowork-ollama ollama pull $Model
if ($LASTEXITCODE -ne 0) { throw "ollama pull failed for $Model" }

$Mode = if ($GpuInContainer) { "GPU ($GpuName)" } elseif ($UseGpu) { "CPU (GPU detected but not passing through)" } else { "CPU" }

$KreuzbergStatus = if ($kzReady) { "ready (http://localhost:8000)" } else { "NOT READY — pdfplumber fallback will be used" }

Write-Host ""
Write-Host "Bootstrap complete."
Write-Host "  Python          : $PythonCmd"
Write-Host "  PaddleOCR image : $PaddleImage"
Write-Host "  Kreuzberg       : $KreuzbergStatus"
Write-Host "  Ollama model    : $Model"
Write-Host "  Ollama endpoint : http://localhost:11434"
Write-Host "  Inference mode  : $Mode"
Write-Host "  Project dir     : $ProjectDir"
Write-Host ""
Write-Host "Next: drop a stock-statement PDF into $(Join-Path $ProjectDir 'inbox') and run:"
Write-Host "    $PythonCmd $(Join-Path $ProjectDir 'worker.py')"
