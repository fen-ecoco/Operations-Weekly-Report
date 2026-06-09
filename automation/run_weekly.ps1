# =============================================================
# ecoco 每週客訴週報自動化主流程
# 每週一 11:00 前自動執行 Step 1~4 + Push + 本機同步
# =============================================================

param(
    [switch]$DryRun   # 加 -DryRun 僅本機執行，不 push
)

$ErrorActionPreference = "Stop"
$BASE = Split-Path -Parent $MyInvocation.MyCommand.Path
$CFG  = Get-Content "$BASE\config.json" -Encoding UTF8 | ConvertFrom-Json
$LOG  = "$BASE\run_log.txt"

$CSV_PATH  = $CFG.csv_path
$OUT_DIR   = $CFG.output_dir          # D:\info\0507_Weekly-Report
$REPO_DIR  = $CFG.repo_dir            # D:\info\0507_Weekly-Report\Operations-Weekly-Report
$PYTHON    = $CFG.python_path
$NODE      = $CFG.node_path
$GIT       = $CFG.git_path
$GH_USER   = $CFG.github_user
$GH_PAT    = $CFG.github_pat
$GH_REPO   = $CFG.github_repo

function Log($msg) {
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content -Path $LOG -Value $line -Encoding UTF8
}

function Die($msg) { Log "❌ $msg"; exit 1 }

Log "=========================================="
Log "  ecoco 週報自動化流程啟動"
Log "=========================================="

# ──────────────────────────────────────────
# Step 1 + 2：載入 CSV → 分析 → data.json
# ──────────────────────────────────────────
Log "[Step 1] 載入 CSV：$CSV_PATH"
if (-not (Test-Path $CSV_PATH)) { Die "CSV 找不到：$CSV_PATH" }

Log "[Step 2] 分析資料，更新 DATA 物件 → data.json"
& $PYTHON "$BASE\analyze.py"
if ($LASTEXITCODE -ne 0) { Die "analyze.py 執行失敗" }

$D      = Get-Content "$BASE\data.json" -Encoding UTF8 | ConvertFrom-Json
$WEEK   = $D.week
$RANGE  = $D.range -replace " ","" -replace "/","" -replace "~","-"
$FNAME  = "ecoco_週報_${WEEK}_${RANGE}.pptx"
$PPTX   = "$BASE\$FNAME"
Log "  ✅ 分析完成：$WEEK（$($D.range)），共 $($D.total) 件已分類"

# ──────────────────────────────────────────
# Step 3：產出 PPT
# ──────────────────────────────────────────
Log "[Step 3] 產出 PPT：$FNAME"
& $NODE "$BASE\generate_ppt_auto.js"
if ($LASTEXITCODE -ne 0) { Die "generate_ppt_auto.js 執行失敗" }
if (-not (Test-Path $PPTX)) { Die "PPTX 未產出：$PPTX" }
Log "  ✅ PPT 產出完成"

# 同步複製到本機輸出目錄
if (-not (Test-Path $OUT_DIR)) { New-Item -ItemType Directory -Path $OUT_DIR | Out-Null }
$LOCAL_COPY = "$OUT_DIR\$FNAME"
Copy-Item $PPTX $LOCAL_COPY -Force
Log "  ✅ 本機同步：$LOCAL_COPY"

# ──────────────────────────────────────────
# Step 4：視覺 QA（開啟 PPTX 目視確認）
# ──────────────────────────────────────────
Log "[Step 4] 視覺 QA — 開啟 PPTX 確認版面"
Start-Process $LOCAL_COPY
Start-Sleep -Seconds 3
Log "  ✅ PPTX 已開啟，請確認版面後繼續"

if ($DryRun) {
    Log "⚠ DryRun 模式：略過 Git push"
    Log "=========================================="
    Log "  流程完成（DryRun）"
    Log "=========================================="
    exit 0
}

# ──────────────────────────────────────────
# Git Commit + Push
# ──────────────────────────────────────────
Log "[Push] Git commit & push → GitHub"
$REPO_PPTX   = "$REPO_DIR\weekly-ppt\$FNAME"
$REPO_HIST   = "$REPO_DIR\automation\history.json"

# 確保 REPO 存在（首次 clone）
if (-not (Test-Path $REPO_DIR)) {
    Log "  首次執行：clone repo..."
    Set-Location (Split-Path $REPO_DIR)
    & $GIT clone "https://${GH_USER}:${GH_PAT}@$($GH_REPO.Replace('https://',''))" | ForEach-Object { Log "  $_" }
}

# 複製 PPTX 與 history 到 repo
if (-not (Test-Path "$REPO_DIR\weekly-ppt")) { New-Item -ItemType Directory "$REPO_DIR\weekly-ppt" | Out-Null }
Copy-Item $PPTX          $REPO_PPTX -Force
Copy-Item "$BASE\history.json" $REPO_HIST -Force

Set-Location $REPO_DIR

& $GIT config user.email "fen-ecoco@ecoco.com.tw"
& $GIT config user.name  "fen-ecoco"
& $GIT add "weekly-ppt\$FNAME" "automation\history.json" 2>&1 | ForEach-Object { Log "  $_" }

$MSG = "feat(${WEEK}): 客訴週報自動產出 $($D.range)`n`n總件數 $($D.total) 件 | 機台 $($D.cats[1].pct)% | APP帳號 $($D.cats[0].pct)%`n自動化流程：analyze.py → generate_ppt_auto.js"
& $GIT commit -m $MSG 2>&1 | ForEach-Object { Log "  $_" }

$REMOTE = "https://${GH_USER}:${GH_PAT}@$($GH_REPO.Replace('https://',''))"
& $GIT push $REMOTE main 2>&1 | ForEach-Object { Log "  $_" }
if ($LASTEXITCODE -ne 0) { Die "Push 失敗（PAT 可能已過期，請更新 config.json 的 github_pat）" }

Log "  ✅ Push 成功 → $GH_REPO"

# ──────────────────────────────────────────
# 完成摘要
# ──────────────────────────────────────────
Log "=========================================="
Log "  ✅ 所有步驟完成"
Log "  PPTX 本機：$LOCAL_COPY"
Log "  PPTX GitHub：weekly-ppt/$FNAME"
Log "  歷史資料：automation/history.json"
Log "=========================================="
