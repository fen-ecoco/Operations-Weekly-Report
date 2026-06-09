# ============================================================
# 一次性設定：Windows 工作排程器
# 每週一 11:00 自動執行 run_weekly.ps1
# 以系統管理員身分執行此腳本
# ============================================================

$TASK_NAME = "ecoco_週報自動化"
$BASE      = Split-Path -Parent $MyInvocation.MyCommand.Path
$SCRIPT    = "$BASE\run_weekly.ps1"
$POWERSHELL = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

# 建立觸發條件：每週一 11:00
$trigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Monday `
    -At "11:00AM"

# 建立動作：執行 PowerShell 腳本
$action = New-ScheduledTaskAction `
    -Execute $POWERSHELL `
    -Argument "-NonInteractive -ExecutionPolicy Bypass -File `"$SCRIPT`"" `
    -WorkingDirectory $BASE

# 設定：即使未登入也執行，最高權限
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

# 登錄工作排程
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType S4U `
    -RunLevel Highest

Register-ScheduledTask `
    -TaskName $TASK_NAME `
    -Trigger $trigger `
    -Action $action `
    -Settings $settings `
    -Principal $principal `
    -Description "ecoco 每週一 11:00 自動執行客訴週報分析、產出 PPT 並推送至 GitHub" `
    -Force

Write-Host "✅ 工作排程已建立：$TASK_NAME"
Write-Host "   執行時間：每週一 11:00"
Write-Host "   腳本路徑：$SCRIPT"
Write-Host ""
Write-Host "確認排程："
Get-ScheduledTask -TaskName $TASK_NAME | Format-List TaskName, State, Triggers
