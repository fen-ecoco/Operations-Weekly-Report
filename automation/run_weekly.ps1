# ecoco Weekly Report - Main Orchestrator (v6, ASCII-only)
# Trigger: Every Monday 11:00 (Task Scheduler), StartWhenAvailable fallback enabled
# Steps: 1) Check input files  2) Run analyze.py  3) Run generate_ppt_auto.js
#        4) Copy output PPTX to save folder  5) Git add/commit/push

# ---------------- Paths ----------------
# NOTE: actual (Chinese-named) file paths live only in config.json (UTF-8),
# never as literal strings in this .ps1 file, to avoid the encoding corruption
# issue documented from earlier runs. Edit config.json, not this script, to
# change source CSV locations or the output folder.
$AutomationDir = "D:\info\0507_Weekly-Report\Operations-Weekly-Report\automation"

$ConfigPath = Join-Path $AutomationDir "config.json"
$Config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$ComplaintCsv = $Config.complaint_csv_path
$VolumeCsv    = $Config.volume_csv_path
$GradeCsv     = $Config.grade_report_path
$OutputDir    = $Config.output_dir

$LogFile = Join-Path $AutomationDir "run_weekly.log"

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $Message" | Out-File -FilePath $LogFile -Append -Encoding utf8
    Write-Host "$ts  $Message"
}

Write-Log "===== Weekly report run started ====="

# ---------------- Step 1: verify input files exist ----------------
if (-not (Test-Path $ComplaintCsv)) {
    Write-Log "ERROR: complaint csv not found at $ComplaintCsv"
    exit 1
}
if (-not (Test-Path $VolumeCsv)) {
    Write-Log "ERROR: volume csv not found at $VolumeCsv"
    exit 1
}
if (-not (Test-Path $GradeCsv)) {
    Write-Log "ERROR: grade report csv (收瓶量分析報告.csv) not found at $GradeCsv"
    exit 1
}
Write-Log "Step 1 OK: all source CSV files found"

# ---------------- Step 2: run analyze.py ----------------
Set-Location $AutomationDir
Write-Log "Step 2: running analyze.py"
python analyze.py
if ($LASTEXITCODE -ne 0) {
    Write-Log "ERROR: analyze.py failed with exit code $LASTEXITCODE"
    exit 1
}
Write-Log "Step 2 OK: data.json generated"

# ---------------- Step 3: run generate_ppt_auto.js ----------------
Write-Log "Step 3: running generate_ppt_auto.js"
node generate_ppt_auto.js
if ($LASTEXITCODE -ne 0) {
    Write-Log "ERROR: generate_ppt_auto.js failed with exit code $LASTEXITCODE"
    exit 1
}
Write-Log "Step 3 OK: PPTX generated"

# ---------------- Step 4: copy output PPTX to save folder ----------------
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
    Write-Log "Created output folder: $OutputDir"
}
$LatestPptx = Get-ChildItem -Path $AutomationDir -Filter "*.pptx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $LatestPptx) {
    Write-Log "ERROR: no pptx file found after generation"
    exit 1
}
Copy-Item -Path $LatestPptx.FullName -Destination $OutputDir -Force
Write-Log "Step 4 OK: copied $($LatestPptx.Name) to $OutputDir"

# ---------------- Step 5: git add / commit / push ----------------
Write-Log "Step 5: git push"
git add .
git commit -m "Weekly report auto-update $(Get-Date -Format 'yyyy-MM-dd')"
git push
if ($LASTEXITCODE -ne 0) {
    Write-Log "WARNING: git push returned non-zero exit code (check remote/auth). PPTX was still saved locally."
} else {
    Write-Log "Step 5 OK: pushed to GitHub"
}

Write-Log "===== Weekly report run finished successfully ====="
