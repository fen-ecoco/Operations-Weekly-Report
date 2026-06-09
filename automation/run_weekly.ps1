# ============================================================
# ecoco 每週客訴自動化流程
# 每週一中午前自動執行 Step 1~4 + Push + 本機同步
# ============================================================

param(
    [switch]$DryRun    # 加 -DryRun 僅測試，不 push
)

$ErrorActionPreference = "Stop"
$BASE    = Split-Path -Parent $MyInvocation.MyCommand.Path
$CFG     = Get-Content "$BASE\config.json" | ConvertFrom-Json
$LOG     = "$BASE\run_log.txt"
$PYTHON  = $CFG.python_path   # e.g. "python" or "C:\Python311\python.exe"
$NODE    = $CFG.node_path     # e.g. "node" or "C:\Program Files\nodejs\node.exe"
$GIT     = $CFG.git_path      # e.g. "git"
$REPO    = $CFG.github_repo
$USER    = $CFG.github_user
$PAT     = $CFG.github_pat
$OUT_DIR = $CFG.output_dir

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line
    Add-Content -Path $LOG -Value $line -Encoding UTF8
}

Log "===== 週報自動化流程開始 ====="

# ── Step 1 + 2：CSV 分析 → data.json ──
Log "[Step 1+2] 分析 CSV → data.json"
& $PYTHON "$BASE\analyze.py"
if ($LASTEXITCODE -ne 0) { Log "❌ analyze.py 失敗"; exit 1 }
$D = Get-Content "$BASE\data.json" | ConvertFrom-Json
$WEEK  = $D.week
$RANGE = $D.range -replace " ", "" -replace "/", "" -replace "~", "-"
$FNAME = "ecoco_週報_${WEEK}_${RANGE}.pptx"
Log "✅ 資料分析完成：$WEEK（$($D.range)），共 $($D.total) 件"

# ── Step 3：產出 PPT ──
Log "[Step 3] 產出 PPT：$FNAME"
& $NODE "$BASE\generate_ppt_auto.js"
if ($LASTEXITCODE -ne 0) { Log "❌ generate_ppt_auto.js 失敗"; exit 1 }
Log "✅ PPT 產出完成"

# ── Step 4：視覺 QA（用 PowerShell 開啟 PPTX 供人工確認）──
Log "[Step 4] 視覺 QA — 開啟 PPTX 確認版面"
$PPTX_PATH = "$BASE\$FNAME"
if (Test-Path $PPTX_PATH) {
    Start-Process $PPTX_PATH
    Start-Sleep -Seconds 5   # 等待開啟
    Log "✅ PPTX 已開啟，請目視確認版面"
} else {
    Log "⚠ 找不到 PPTX：$PPTX_PATH"
}

if ($DryRun) {
    Log "⚠ DryRun 模式：略過 Git push"
    Log "===== 完成（DryRun）====="
    exit 0
}

# ── Git Commit + Push ──
Log "[Push] Git commit & push → GitHub"
$REPO_ROOT = Split-Path -Parent $BASE
Set-Location $REPO_ROOT

& $GIT add "weekly-ppt\$FNAME" 2>&1 | ForEach-Object { Log $_ }
& $GIT add "automation\history.json" 2>&1 | ForEach-Object { Log $_ }

$MSG = "feat($WEEK): 客訴週報自動產出 $($D.range)`n`n- 總件數 $($D.total) 件（已分類）`n- 機台問題 $($D.cats[1].pct)%　APP帳號設定 $($D.cats[0].pct)%`n- 自動化流程：analyze.py → generate_ppt_auto.js"
& $GIT commit -m $MSG 2>&1 | ForEach-Object { Log $_ }

$REMOTE = "https://${USER}:${PAT}@github.com/$($REPO.Split('/')[-2])/$($REPO.Split('/')[-1])"
& $GIT push $REMOTE main 2>&1 | ForEach-Object { Log $_ }
if ($LASTEXITCODE -ne 0) { Log "❌ Git push 失敗（PAT 可能已過期，請更新 config.json）"; exit 1 }

Log "✅ Push 成功 → $REPO"

# ── 本機同步確認 ──
$LOCAL_PPTX = "$OUT_DIR\$FNAME"
if (Test-Path $LOCAL_PPTX) {
    Log "✅ 本機已同步：$LOCAL_PPTX"
} else {
    Log "⚠ 本機檔案未找到，嘗試複製..."
    Copy-Item $PPTX_PATH $OUT_DIR -Force
    Log "✅ 複製完成：$LOCAL_PPTX"
}

Log "===== 週報自動化流程完成 ✅ ====="
