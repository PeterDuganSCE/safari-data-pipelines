$ErrorActionPreference = "Stop"

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$logFile = "run_log.txt"

function Log {
    param([string]$message)
    $entry = "$((Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) - $message"
    Write-Host $entry
    Add-Content -Path $logFile -Value $entry
}

try {
    Log "Starting job"

    Set-Location -Path $PSScriptRoot

    if (Test-Path ".\.venv\Scripts\Activate.ps1") {
        Log "Activating virtual environment"
        . .\.venv\Scripts\Activate.ps1
    }

    Log "Running Python script"
    python main.py

    Log "Job completed successfully"
}
catch {
    Log "ERROR: $($_.Exception.Message)"
    exit 1
}