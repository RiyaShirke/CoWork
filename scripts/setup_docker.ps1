# One-time Docker bring-up for the Cowork worker.
# Run this when you are ready to actually process bills.
#   powershell -ExecutionPolicy Bypass -File D:\Cowork\scripts\setup_docker.ps1
#
# What it does:
#   1. Pulls paddlecloud/paddleocr (used per-file via `docker run --rm`).
#   2. Starts the Ollama container in the background via docker-compose.
#   3. Pulls the llama3.2:1b model inside the Ollama container.
#
# Override the model by setting $env:OLLAMA_MODEL before running, e.g.:
#   $env:OLLAMA_MODEL = 'llama3.2:3b'; .\setup_docker.ps1

$ErrorActionPreference = 'Stop'
$ProjectDir = 'C:\Cowork\Cowork'
$Model = if ($env:OLLAMA_MODEL) { $env:OLLAMA_MODEL } else { 'llama3.2:1b' }
$PaddleImage = if ($env:PADDLE_IMAGE) { $env:PADDLE_IMAGE } else { 'paddlecloud/paddleocr:2.6-cpu-latest' }

Write-Host "==> Pulling PaddleOCR image: $PaddleImage"
docker pull $PaddleImage
if ($LASTEXITCODE -ne 0) { throw "docker pull failed for $PaddleImage" }

Write-Host ""
Write-Host "==> Bringing up Ollama (docker compose)"
docker compose -f (Join-Path $ProjectDir 'docker-compose.yml') up -d
if ($LASTEXITCODE -ne 0) { throw "docker compose up failed" }

Write-Host ""
Write-Host "==> Waiting for Ollama to accept connections on :11434"
for ($i = 0; $i -lt 30; $i++) {
    try {
        $r = Invoke-WebRequest -Uri 'http://localhost:11434/api/tags' -UseBasicParsing -TimeoutSec 2
        if ($r.StatusCode -eq 200) { Write-Host "    ready."; break }
    } catch { Start-Sleep -Seconds 2 }
}

Write-Host ""
Write-Host "==> Pulling Ollama model: $Model (this may take a few minutes)"
docker exec cowork-ollama ollama pull $Model
if ($LASTEXITCODE -ne 0) { throw "ollama pull failed for $Model" }

Write-Host ""
Write-Host "Docker setup complete."
Write-Host "  PaddleOCR image : $PaddleImage"
Write-Host "  Ollama model    : $Model"
Write-Host "  Ollama endpoint : http://localhost:11434"
