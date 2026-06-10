# ecoco Weekly Report Automation
# Run every Monday at 11:00 AM (via Task Scheduler)
# Steps: 1+2 Analyze CSV -> 3 Generate PPT -> 4 QA -> Git Push -> Local Sync

param([switch]$DryRun)

$ErrorActionPreference = "Stop"
$BASE   = Split-Path -Parent $MyInvocation.MyCommand.Path
$CFG    = Get-Content "$BASE\config.json" -Encoding UTF8 | ConvertFrom-Json
$LOG    = "$BASE\run_log.txt"

$CSV_PATH = $CFG.csv_path
$OUT_DIR  = $CFG.output_dir
$REPO_DIR = $CFG.repo_dir
$PYTHON   = $CFG.python_path
$NODE     = $CFG.node_path
$GIT      = $CFG.git_path
$GH_USER  = $CFG.github_user
$GH_PAT   = $CFG.github_pat
$GH_REPO  = $CFG.github_repo

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content -Path $LOG -Value $line -Encoding UTF8
}

Log "=========================================="
Log "  ecoco Weekly Report - Start"
Log "=========================================="

# --- Step 1+2: Analyze CSV -> data.json ---
Log "[Step 1+2] Load CSV and analyze..."
if (-not (Test-Path $CSV_PATH)) {
    Log "ERROR: CSV not found: $CSV_PATH"
    exit 1
}
& $PYTHON "$BASE\analyze.py"
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: analyze.py failed"
    exit 1
}
$D      = Get-Content "$BASE\data.json" -Encoding UTF8 | ConvertFrom-Json
$WEEK   = $D.week
$RCLEAN = $D.range -replace " ","" -replace "/","" -replace "~","-"
$FNAME  = "ecoco_weekly_${WEEK}_${RCLEAN}.pptx"
$PPTX   = "$BASE\$FNAME"
Log "  OK: $WEEK ($($D.range)) - $($D.total) records"

# --- Step 3: Generate PPT ---
Log "[Step 3] Generate PPT: $FNAME"
& $NODE "$BASE\generate_ppt_auto.js"
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: generate_ppt_auto.js failed"
    exit 1
}
if (-not (Test-Path $PPTX)) {
    Log "ERROR: PPTX not found after generation: $PPTX"
    exit 1
}
Log "  OK: PPT generated"

# Copy to output dir (local sync)
if (-not (Test-Path $OUT_DIR)) {
    New-Item -ItemType Directory -Path $OUT_DIR | Out-Null
}
$LOCAL_COPY = "$OUT_DIR\$FNAME"
Copy-Item $PPTX $LOCAL_COPY -Force
Log "  OK: Local sync -> $LOCAL_COPY"

# --- Step 4: Visual QA ---
Log "[Step 4] Opening PPTX for visual check..."
Start-Process $LOCAL_COPY
Start-Sleep -Seconds 3
Log "  OK: PPTX opened - please verify layout"

if ($DryRun) {
    Log "DryRun mode: skip Git push"
    Log "=========================================="
    Log "  ecoco Weekly Report - Done (DryRun)"
    Log "=========================================="
    exit 0
}

# --- Git Push ---
Log "[Push] Git commit and push..."

$REPO_PPTX = "$REPO_DIR\weekly-ppt\$FNAME"
$REPO_HIST = "$REPO_DIR\automation\history.json"

if (-not (Test-Path $REPO_DIR)) {
    Log "  First run: cloning repo..."
    Set-Location (Split-Path $REPO_DIR)
    $GH_URL = $GH_REPO -replace "https://", ""
    & $GIT clone "https://${GH_USER}:${GH_PAT}@${GH_URL}" | ForEach-Object { Log "  $_" }
}

if (-not (Test-Path "$REPO_DIR\weekly-ppt")) {
    New-Item -ItemType Directory "$REPO_DIR\weekly-ppt" | Out-Null
}
Copy-Item $PPTX          $REPO_PPTX -Force
Copy-Item "$BASE\history.json" $REPO_HIST -Force

Set-Location $REPO_DIR
& $GIT config user.email "fen-ecoco@ecoco.com.tw"
& $GIT config user.name "fen-ecoco"
& $GIT add "weekly-ppt\$FNAME" "automation\history.json" 2>&1 | ForEach-Object { Log "  $_" }

$MSG = "feat($WEEK): weekly report $($D.range) - $($D.total) records"
& $GIT commit -m $MSG 2>&1 | ForEach-Object { Log "  $_" }

$GH_URL = $GH_REPO -replace "https://", ""
& $GIT push "https://${GH_USER}:${GH_PAT}@${GH_URL}" main 2>&1 | ForEach-Object { Log "  $_" }
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: Push failed - check github_pat in config.json"
    exit 1
}
Log "  OK: Pushed to GitHub -> $GH_REPO"

# --- Summary ---
Log "=========================================="
Log "  OK: All steps completed"
Log "  PPTX local : $LOCAL_COPY"
Log "  PPTX GitHub: weekly-ppt/$FNAME"
Log "  History    : automation/history.json"
Log "=========================================="
