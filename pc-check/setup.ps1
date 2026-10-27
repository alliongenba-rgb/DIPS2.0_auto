# 成果物 格納チェックの初期設定（ALLION-01 で1回だけ実行する）。
# 「setup.bat」をダブルクリックすると、このスクリプトが動く。
#  1. 一式を C:\pc-check にコピー
#  2. 成果物の場所（N:\01_社内共有\01_案件データ\2026年度\...）を確認
#  3. Notion の連携用の鍵を登録し、読み書きできるか確認
#  4. config.json を作成し、自動実行を登録し、1回目のチェックを行う
$ErrorActionPreference = 'Stop'

$installDir = 'C:\pc-check'
$defaultTarget = 'N:\01_社内共有\01_案件データ'
$fromYear = 2026
$yearPattern = '^\s*(\d{4})\s*年度'
$taskName = 'PC-Check-Collect'
$logDataSourceId = 'a8600581-6f48-4b75-b1ec-9cafa9713dac'     # 📱 作業経過ログ（外出先確認用）
$checkDataSourceId = '233e9b89-0f24-40f8-a494-6fb9739a63e3'   # 📦 成果物 格納チェック（自動更新）

function Write-Step($n, $text) {
    Write-Host ''
    Write-Host "[$n/4] $text" -ForegroundColor Cyan
}

# 割り当てたドライブ（N: など）は自動実行のときに見えないことがあるため、\\サーバー\共有 の形に直す
function ConvertTo-UncPath($path) {
    $full = (Resolve-Path -LiteralPath $path).ProviderPath
    if ($full -match '^([A-Za-z]):') {
        $drive = Get-PSDrive -Name $Matches[1] -PSProvider FileSystem -ErrorAction SilentlyContinue
        if ($drive -and $drive.DisplayRoot -like '\\*') {
            return $drive.DisplayRoot.TrimEnd('\') + $full.Substring(2)
        }
    }
    return $full
}

# Notion の2つのデータベースが、この鍵で開けるか確かめる。開けなければ理由を返す
function Test-NotionAccess($token) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $headers = @{ Authorization = "Bearer $token"; 'Notion-Version' = '2025-09-03' }
    foreach ($ds in @(@{ Id = $logDataSourceId; Name = '📱 作業経過ログ' }, @{ Id = $checkDataSourceId; Name = '📦 成果物 格納チェック' })) {
        try {
            Invoke-RestMethod -Uri "https://api.notion.com/v1/data_sources/$($ds.Id)" -Headers $headers -Method Get | Out-Null
        } catch {
            $status = 0
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if ($status -eq 401) { return '鍵が正しくありません（コピーし直してください）' }
            if ($status -eq 404) { return "「$($ds.Name)」を開く権限がありません（手順の「アクセス権」を確認してください）" }
            return "Notion に接続できません: $($_.Exception.Message)"
        }
    }
    return $null
}

Write-Host '==============================================' -ForegroundColor Green
Write-Host '  成果物 格納チェック  セットアップ' -ForegroundColor Green
Write-Host '==============================================' -ForegroundColor Green

# 実行前に、作成・登録するものを示して承認を得る（共通行動規範 RULE 3）
Write-Host ''
Write-Host "  実行するPC          : $env:COMPUTERNAME"
Write-Host "  作成・上書きするもの: $installDir （プログラム置き場）"
Write-Host "  登録するタスク      : $taskName （30分ごとに格納チェック。前回の同名タスクは置き換え）"
Write-Host '  保存する鍵          : Notion の連携用の鍵（このPCのユーザー環境変数 ALLION_NOTION_TOKEN）'
Write-Host "  読むだけの場所      : $defaultTarget\${fromYear}年度 （変更・削除はしない）"
Write-Host '  書き込む場所        : Notion「📦 成果物 格納チェック（自動更新）」'
Write-Host ''
if ((Read-Host '  この内容で進めてよければ Enter、やめる場合は n を入力') -eq 'n') {
    Write-Host '  中止しました。何も変更していません。'
    return
}

# --- 1. コピー ---
Write-Step 1 "プログラムを $installDir にコピーします"
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
if ($PSScriptRoot.TrimEnd('\') -ne $installDir) {
    foreach ($name in @('check.ps1', 'install.ps1', 'setup.ps1', 'setup.bat', 'README.md')) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $installDir -Force
    }
}
Get-ChildItem -LiteralPath $installDir -File | Unblock-File
Write-Host '  完了'

# --- 2. 成果物の場所を確認 ---
Write-Step 2 "成果物の場所を確認します（${fromYear}年度フォルダが入っているフォルダ）"
$target = $null
while (-not $target) {
    $ans = Read-Host "  ${fromYear}年度フォルダが入っているフォルダ（そのままEnterで $defaultTarget）"
    if (-not $ans) { $ans = $defaultTarget }
    $ans = $ans.Trim('"')
    if (-not (Test-Path -LiteralPath $ans)) {
        Write-Host "  $ans が見つかりません。ドライブがつながっているか確認してください。" -ForegroundColor Yellow
        continue
    }
    # 年度フォルダと、その直下のフォルダを表示して、場所が合っているか確認してもらう
    $years = @(Get-ChildItem -LiteralPath $ans -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $yearPattern -and [int]$Matches[1] -ge $fromYear })
    if ($years.Count -eq 0) {
        Write-Host "  $ans の中に ${fromYear}年度 のフォルダが見つかりません。" -ForegroundColor Yellow
        continue
    }
    foreach ($y in $years) {
        $subs = @(Get-ChildItem -LiteralPath $y.FullName -Directory -ErrorAction SilentlyContinue)
        Write-Host ''
        Write-Host "  $($y.Name) の直下のフォルダ: $($subs.Count) 件"
        $subs | Select-Object -First 10 | ForEach-Object { Write-Host "    $($_.Name)" }
        if ($subs.Count -gt 10) { Write-Host "    ...ほか $($subs.Count - 10) 件" }
    }
    Write-Host ''
    $ok = Read-Host '  この場所でよければ Enter、やり直す場合は n を入力'
    if ($ok -ne 'n') { $display = $ans; $target = ConvertTo-UncPath $ans }
}
Write-Host "  成果物の場所: $target" -ForegroundColor Green

# --- 3. Notion の鍵 ---
Write-Step 3 'Notion の連携用の鍵を登録します'
$token = [Environment]::GetEnvironmentVariable('ALLION_NOTION_TOKEN', 'User')
if ($token) {
    if ((Read-Host '  登録済みの鍵があります。そのまま使うなら Enter、入れ直すなら n を入力') -eq 'n') { $token = $null }
}
while ($true) {
    if (-not $token) {
        $secure = Read-Host '  Notion でコピーした鍵（ntn_ で始まる文字）を貼り付けて Enter（画面には表示されません）' -AsSecureString
        $token = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)).Trim()
    }
    $problem = Test-NotionAccess $token
    if (-not $problem) { break }
    Write-Host "  $problem" -ForegroundColor Yellow
    if ((Read-Host '  直したら Enter でもう一度確かめます（中止は n）') -eq 'n') {
        Write-Host '  中止しました。プログラムのコピーだけ済んでいます。'
        return
    }
    if ($problem.StartsWith('鍵が正しく')) { $token = $null }
}
[Environment]::SetEnvironmentVariable('ALLION_NOTION_TOKEN', $token, 'User')
Write-Host '  鍵を確認し、このPCに保存しました（Notion の2つのデータベースを開けることを確認）' -ForegroundColor Green

# --- 4. 設定・自動実行・1回目のチェック ---
Write-Step 4 '設定を保存し、自動実行を登録して、1回目のチェックを行います（数分かかります）'
$config = [ordered]@{
    target            = $target
    displayRoot       = $display
    fromYear          = $fromYear
    excludeNames      = @('~$*', '*.tmp', '*.bak', 'desktop.ini', 'Thumbs.db', '.DS_Store')
    excludeDirs       = @('*.files')
    logDataSourceId   = $logDataSourceId
    checkDataSourceId = $checkDataSourceId
    logDays           = 14
}
$config | ConvertTo-Json | Set-Content -Path (Join-Path $installDir 'config.json') -Encoding UTF8
& (Join-Path $installDir 'install.ps1') -TaskName $taskName
$env:ALLION_NOTION_TOKEN = $token
& (Join-Path $installDir 'check.ps1')

Write-Host ''
Write-Host '==============================================' -ForegroundColor Green
Write-Host '  セットアップ完了' -ForegroundColor Green
Write-Host '  以後30分ごとに自動でチェックし、Notion に書き込みます。' -ForegroundColor Green
Write-Host '  確認は Claude に「今日の作業、所定の場所に格納されてる？」と聞くだけです。' -ForegroundColor Green
Write-Host '==============================================' -ForegroundColor Green
