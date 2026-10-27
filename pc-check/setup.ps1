# PC成果物チェックの初期設定（各PCで1回だけ実行する）。
# 「setup.bat」をダブルクリックすると、このスクリプトが動く。
#  1. 一式を C:\pc-check にコピー
#  2. 成果物の場所（N:\01_社内共有\01_案件データ\2026年度\...）を確認
#  3. config.json を作成し、自動実行を登録し、1回目の収集を行う
$ErrorActionPreference = 'Stop'

$installDir = 'C:\pc-check'
$defaultTarget = 'N:\01_社内共有\01_案件データ'
$defaultOutputDir = 'C:\pc-check-share'
$fromYear = 2026
$yearPattern = '^\s*(\d{4})\s*年度'

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

Write-Host '==============================================' -ForegroundColor Green
Write-Host '  PC成果物チェック  セットアップ' -ForegroundColor Green
Write-Host '==============================================' -ForegroundColor Green

# 実行前に、作成・登録するものを示して承認を得る（共通行動規範 RULE 3）
Write-Host ''
Write-Host "  実行するPC          : $env:COMPUTERNAME"
Write-Host "  作成するフォルダ    : $installDir （プログラム置き場）"
Write-Host "                        $defaultOutputDir （一覧の出力先・変更可）"
Write-Host '  登録するタスク      : PC-Check-Collect （30分ごとに一覧を更新）'
Write-Host "  読むだけの場所      : $defaultTarget\${fromYear}年度 （変更・削除はしない）"
Write-Host ''
if ((Read-Host '  この内容で進めてよければ Enter、やめる場合は n を入力') -eq 'n') {
    Write-Host '  中止しました。何も変更していません。'
    return
}

# --- 1. コピー ---
Write-Step 1 "プログラムを $installDir にコピーします"
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
if ($PSScriptRoot.TrimEnd('\') -ne $installDir) {
    foreach ($name in @('collect.ps1', 'install.ps1', 'setup.ps1', 'setup.bat', 'viewer.html', 'README.md')) {
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
    if ($ok -ne 'n') { $target = ConvertTo-UncPath $ans }
}
Write-Host "  成果物の場所: $target" -ForegroundColor Green

# --- 3. 設定ファイルを作る ---
Write-Step 3 '設定を保存します'
$out = Read-Host "  一覧の出力先（そのままEnterで $defaultOutputDir）"
if (-not $out) { $out = $defaultOutputDir }
$config = [ordered]@{
    targets      = @($target)
    fromYear     = $fromYear
    extensions   = @()
    excludeNames = @('~$*', '*.tmp', '*.bak', 'desktop.ini', 'Thumbs.db', '.DS_Store')
    excludeDirs  = @('*.files')
    outputDir    = $out.Trim('"')
}
$config | ConvertTo-Json | Set-Content -Path (Join-Path $installDir 'config.json') -Encoding UTF8
Write-Host "  保存しました: $installDir\config.json"

# --- 4. 自動実行の登録と1回目の収集 ---
Write-Step 4 '自動実行を登録し、1回目の収集を行います（件数が多いと数分かかります）'
& (Join-Path $installDir 'install.ps1')
& (Join-Path $installDir 'collect.ps1')

# 何の種類のファイルが集まったかを表示する（各ソフトの出力が含まれているかの確認用）
$csv = Join-Path $config.outputDir "$env:COMPUTERNAME.csv"
$rows = @(Import-Csv -LiteralPath $csv -Encoding UTF8 | Where-Object { $_.Name })
if ($rows.Count -gt 0) {
    Write-Host ''
    Write-Host '  集まったファイルの種類（多い順・上位15）:'
    $rows | Group-Object Extension | Sort-Object Count -Descending | Select-Object -First 15 |
        ForEach-Object { Write-Host ('    {0,-10} {1,7} 件' -f $(if ($_.Name) { $_.Name } else { '(拡張子なし)' }), $_.Count) }
}

Write-Host ''
Write-Host '==============================================' -ForegroundColor Green
Write-Host '  セットアップ完了' -ForegroundColor Green
Write-Host "  確認画面: $installDir\viewer.html を開き、" -ForegroundColor Green
Write-Host "  「共有フォルダを選ぶ」で $($config.outputDir) を選択してください。" -ForegroundColor Green
Write-Host '==============================================' -ForegroundColor Green
Start-Process (Join-Path $installDir 'viewer.html')
