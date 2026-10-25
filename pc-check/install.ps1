# collect.ps1 を30分ごとに自動実行するタスクを登録する。
# 各PCで1回だけ実行する（管理者権限は不要。ログイン中のユーザーで動く）。
param(
    [int]$IntervalMinutes = 30,
    [string]$TaskName = 'PC-Check-Collect'
)

$script = Join-Path $PSScriptRoot 'collect.ps1'
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""

$repeat = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$logon = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger @($repeat, $logon) `
    -Settings $settings -Description '成果物ファイル一覧を共有フォルダへ書き出す' -Force | Out-Null

Write-Output "タスク '$TaskName' を登録しました（$IntervalMinutes 分ごと＋ログオン時）。"
Write-Output "今すぐ1回実行して動作確認します..."
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script
