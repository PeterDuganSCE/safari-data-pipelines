param(
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
$logFile = Join-Path $repoRoot "logs\powershell_run_log.txt"
$steps = @(
    @{ Name = "Outlook to raw"; Script = "s1_email_to_safari.py"; Module = "pipelines.watchoffice.s1_email_to_safari" },
    @{ Name = "Raw to staging"; Script = "s2_raw_to_stage.py"; Module = "pipelines.watchoffice.s2_raw_to_stage" },
    @{ Name = "Stage to scrape"; Script = "s3_scrape_wo.py"; Module = "pipelines.watchoffice.s3_scrape_wo" }
)

function Write-Log {
    param([string]$Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - WatchOffice pipeline - $Message"
    Write-Host $entry
    Add-Content -Path $logFile -Value $entry
}

try {
    if (-not (Test-Path $python -PathType Leaf)) {
        throw "Virtual environment Python not found: $python"
    }

    foreach ($step in $steps) {
        $scriptPath = Join-Path $repoRoot "pipelines\watchoffice\$($step.Script)"
        if (-not (Test-Path $scriptPath -PathType Leaf)) {
            throw "Python script not found: $scriptPath"
        }
    }

    Set-Location -Path $repoRoot
    Write-Log "Starting job"

    if ($DryRun) {
        & $python --version
        if ($LASTEXITCODE -ne 0) {
            throw "Python validation failed with exit code $LASTEXITCODE"
        }
        Write-Log "Dry run complete; no pipeline steps were executed"
        exit 0
    }

    foreach ($step in $steps) {
        Write-Log "Starting $($step.Name)"
        & $python -m $step.Module
        if ($LASTEXITCODE -ne 0) {
            throw "$($step.Name) failed with exit code $LASTEXITCODE"
        }
        Write-Log "Completed $($step.Name)"
    }

    Write-Log "Job completed successfully"
    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}