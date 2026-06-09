# =============================================================
# 一次性安裝腳本
# 1. 安裝 Python 套件、Node 套件
# 2. 設定 Windows 工作排程器（每週一 11:00）
# 以系統管理員身分執行
# =============================================================

$ErrorActionPreference = "Stop"
$BASE = Split-Path -Parent $MyInvocation.MyCommand.Path
$CFG  = Get-Content "$BASE\config.json" -Encoding UTF8 | ConvertFrom-Json

Write-Host "=== ecoco 週報自動化 安裝程式 ===" -ForegroundColor Cyan

# ── 1. 設定 config.json ──
if (-not (Test-Path "$BASE\config.json")) {
    Copy-Item "$BASE\config.json.example" "$BASE\config.json"
    Write-Host "⚠ 已建立 config.json，請填入 github_pat 後重新執行" -ForegroundColor Yellow
    Start-Process notepad "$BASE\config.json"
    exit 0
}

# ── 2. 安裝 Python 套件 ──
Write-Host "`n[1/4] 安裝 Python 套件..."
& $CFG.python_path -m pip install pandas --quiet
Write-Host "  ✅ pandas 已安裝"

# ── 3. 安裝 Node 套件 ──
Write-Host "`n[2/4] 安裝 Node 套件..."
Set-Location $BASE
& $CFG.node_path (& $CFG.node_path -e "require('path').join(require('path').dirname(process.execPath),'npm')" 2>$null || "npm") install pptxgenjs --save 2>&1 | Out-Null
# 備援方式
if (-not (Test-Path "$BASE\node_modules\pptxgenjs")) {
    Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c cd `"$BASE`" && npm install pptxgenjs"
}
Write-Host "  ✅ pptxgenjs 已安裝"

# ── 4. 建立工作排程 ──
Write-Host "`n[3/4] 建立 Windows 工作排程器..."
$TASK   = "ecoco_週報自動化"
$PS     = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$SCRIPT = "$BASE\run_weekly.ps1"

$trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At "11:00AM"
$action   = New-ScheduledTaskAction `
    -Execute $PS `
    -Argument "-NonInteractive -ExecutionPolicy Bypass -File `"$SCRIPT`"" `
    -WorkingDirectory $BASE
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType S4U -RunLevel Highest

Register-ScheduledTask `
    -TaskName $TASK -Trigger $trigger `
    -Action $action -Settings $settings `
    -Principal $principal `
    -Description "ecoco 每週一 11:00 自動執行客訴週報分析 → PPT → GitHub" `
    -Force | Out-Null

Write-Host "  ✅ 工作排程已建立：$TASK（每週一 11:00）"

# ── 5. 確認 Git remote ──
Write-Host "`n[4/4] 確認 Git 設定..."
$REPO_DIR = $CFG.repo_dir
if (-not (Test-Path $REPO_DIR)) {
    Write-Host "  首次 clone repo..."
    $GH = $CFG.github_repo.Replace("https://","")
    & $CFG.git_path clone "https://$($CFG.github_user):$($CFG.github_pat)@$GH" $REPO_DIR 2>&1
    Write-Host "  ✅ Repo 已 clone 至：$REPO_DIR"
} else {
    Write-Host "  ✅ Repo 已存在：$REPO_DIR"
}

Write-Host "`n=== 安裝完成 ===" -ForegroundColor Green
Write-Host "排程時間 ：每週一 11:00"
Write-Host "CSV 來源 ：$($CFG.csv_path)"
Write-Host "PPT 輸出 ：$($CFG.output_dir)"
Write-Host "GitHub   ：$($CFG.github_repo)"
Write-Host ""
Write-Host "手動測試（不 push）：" -ForegroundColor Yellow
Write-Host "  .\run_weekly.ps1 -DryRun"
Write-Host "手動完整執行："
Write-Host "  .\run_weekly.ps1"
