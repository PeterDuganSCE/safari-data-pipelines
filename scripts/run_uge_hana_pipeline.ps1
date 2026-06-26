param(
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$scriptRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent $scriptRoot
$logFile = Join-Path $repoRoot "logs/powershell_run_log.txt"

function Log {
    param([string]$message)
    $entry = "$((Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) - $message"
    Write-Host $entry
    Add-Content -Path $logFile -Value $entry
}

try {
    Log "Starting job"
    Log "Initial working directory: $((Get-Location).Path)"
    Log "Script directory: $scriptRoot"
    Log "Repository root: $repoRoot"

    Set-Location -Path $repoRoot
    Log "Execution working directory: $((Get-Location).Path)"

    $activateScript = Join-Path $repoRoot ".venv\Scripts\Activate.ps1"
    $venvPython = Join-Path $repoRoot ".venv\Scripts\python.exe"

    if (-not (Test-Path $activateScript)) {
        throw "Virtual environment activation script not found: $activateScript"
    }

    if (-not (Test-Path $venvPython)) {
        throw "Virtual environment Python executable not found: $venvPython"
    }

    Log "Activating virtual environment: $activateScript"
    try {
        . $activateScript
    }
    catch {
        throw "Virtual environment activation failed: $($_.Exception.Message)"
    }

    if ($DryRun) {
        Log "Dry run enabled. Skipping ETL execution."
        Log "Validating virtual environment Python executable."
        & $venvPython --version
        if ($LASTEXITCODE -ne 0) {
            throw "Dry run validation failed with exit code $LASTEXITCODE"
        }
    }
    else {
        Log "Running Python script"
        & $venvPython -m pipelines.hana.safari_to_hana_uge
        if ($LASTEXITCODE -ne 0) {
            throw "Pipeline execution failed with exit code $LASTEXITCODE"
        }
    }

    Log "Job completed successfully"
}
catch {
    Log "ERROR: $($_.Exception.Message)"
    exit 1
}