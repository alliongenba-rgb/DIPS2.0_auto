# check.ps1（成果物の格納チェック）を定期実行するタスクを登録する（setup.ps1 から呼ばれる）。
# 管理者権限は不要。ログイン中のユーザーで動く。
param(
    [int]$IntervalMinutes = 30,
    [string]$TaskName = 'PC-Check-Collect'
)

$script = Join-Path $PSScriptRoot 'check.ps1'
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""

# 電源オフ等で実行できなかった回は、次に起動したときに実行する（StartWhenAvailable）
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Description 'N: の成果物の格納状況を照合し、Notion に書き込む' -Force | Out-Null

Write-Output "自動実行を登録しました（タスク名: $TaskName、$IntervalMinutes 分ごと）。"
