param(
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
$pipeline = Join-Path $repoRoot "pipelines\repair_order\ro_pdf_scraper.py"
$logFile = Join-Path $repoRoot "logs\powershell_run_log.txt"

function Write-Log {
    param([string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Repair-order PDF scraper - $Message"
    Write-Host $entry
    Add-Content -Path $logFile -Value $entry
}

try {
    if (-not (Test-Path $python -PathType Leaf)) {
        throw "Virtual environment Python not found: $python"
    }

    if (-not (Test-Path $pipeline -PathType Leaf)) {
        throw "Pipeline script not found: $pipeline"
    }

    Set-Location -Path $repoRoot
    Write-Log "Starting job"

    if ($DryRun) {
        & $python --version
        if ($LASTEXITCODE -ne 0) {
            throw "Python validation failed with exit code $LASTEXITCODE"
        }
        Write-Log "Dry run complete; pipeline was not executed"
        exit 0
    }

    Write-Log "Running repair-order PDF scraper"
    & $python -m pipelines.repair_order.ro_pdf_scraper
    if ($LASTEXITCODE -ne 0) {
        throw "Pipeline execution failed with exit code $LASTEXITCODE"
    }

    Write-Log "Job completed successfully"
    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}