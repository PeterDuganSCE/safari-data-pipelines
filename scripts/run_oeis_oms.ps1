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
	$targetScript = Join-Path $repoRoot "reports\oeis_oms.py"
	$targetModule = "reports.oeis_oms"

	if (-not (Test-Path $activateScript)) {
		throw "Virtual environment activation script not found: $activateScript"
	}

	if (-not (Test-Path $venvPython)) {
		throw "Virtual environment Python executable not found: $venvPython"
	}

	if (-not (Test-Path $targetScript)) {
		throw "Target Python script not found: $targetScript"
	}

	Log "Activating virtual environment: $activateScript"
	try {
		. $activateScript
	}
	catch {
		throw "Virtual environment activation failed: $($_.Exception.Message)"
	}

	if ($DryRun) {
		Log "Dry run enabled. Skipping report execution."
		Log "Validating virtual environment Python executable."
		& $venvPython --version
		if ($LASTEXITCODE -ne 0) {
			throw "Dry run validation failed with exit code $LASTEXITCODE"
		}
	}
	else {
		Log "Running Python module: $targetModule"
		& $venvPython -m $targetModule
		if ($LASTEXITCODE -ne 0) {
			throw "Script execution failed with exit code $LASTEXITCODE"
		}
	}

	Log "Job completed successfully"
}
catch {
	Log "ERROR: $($_.Exception.Message)"
	exit 1
}
