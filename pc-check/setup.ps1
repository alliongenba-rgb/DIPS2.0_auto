# PC成果物チェックの初期設定（各PCで1回だけ実行する）。
# 「setup.bat」をダブルクリックすると、このスクリプトが動く。
#  1. 一式を C:\pc-check にコピー
#  2. 「NNNN年度」フォルダを自動で探し、成果物の場所を選んでもらう
#  3. config.json を作成し、自動実行を登録し、1回目の収集を行う
$ErrorActionPreference = 'Stop'

$installDir = 'C:\pc-check'
$defaultOutputDir = 'C:\pc-check-share'
$fromYear = 2006
$yearPattern = '^\s*(\d{4})\s*年度'
$searchDepth = 4
$skipNames = @('Windows', 'Program Files', 'Program Files (x86)', 'ProgramData', 'AppData',
    '$Recycle.Bin', 'System Volume Information', 'Recovery', 'PerfLogs', 'node_modules', '.git')

function Write-Step($n, $text) {
    Write-Host ''
    Write-Host "[$n/4] $text" -ForegroundColor Cyan
}

# 年度フォルダを含む親フォルダを探す（深さ制限つき）
function Find-YearRoots {
    $starts = New-Object System.Collections.Generic.List[string]
    # ネットワークドライブは時間がかかるので、PC内蔵のドライブだけを探す
    [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady } |
        ForEach-Object { $starts.Add($_.RootDirectory.FullName) }
    foreach ($p in @("$env:USERPROFILE\Desktop", "$env:USERPROFILE\Documents", $env:OneDrive)) {
        if ($p -and (Test-Path -LiteralPath $p)) { $starts.Add($p) }
    }

    $found = @{}
    $seen = @{}
    foreach ($start in $starts) {
        $queue = New-Object System.Collections.Generic.Queue[object]
        $queue.Enqueue(@($start, 0))
        while ($queue.Count -gt 0) {
            $entry = $queue.Dequeue()
            $dir = $entry[0]; $depth = $entry[1]
            if ($seen.ContainsKey($dir)) { continue }
            $seen[$dir] = $true
            $children = @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue)
            $years = @($children | Where-Object { $_.Name -match $yearPattern } | ForEach-Object { [int]$Matches[1] })
            if ($years.Count -gt 0) {
                $found[$dir] = $years
                continue
            }
            if ($depth -ge $searchDepth) { continue }
            foreach ($c in $children) {
                if ($skipNames -contains $c.Name -or $c.Name.StartsWith('.')) { continue }
                $queue.Enqueue(@($c.FullName, ($depth + 1)))
            }
        }
    }
    return $found
}

Write-Host '==============================================' -ForegroundColor Green
Write-Host '  PC成果物チェック  セットアップ' -ForegroundColor Green
Write-Host '==============================================' -ForegroundColor Green

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

# --- 2. 成果物フォルダを探す ---
Write-Step 2 '「○○○○年度」フォルダを探しています（1～2分かかることがあります）...'
$found = Find-YearRoots
$candidates = @($found.Keys | Sort-Object)

$target = $null
if ($candidates.Count -gt 0) {
    Write-Host ''
    Write-Host '  次の場所に年度フォルダが見つかりました:'
    for ($i = 0; $i -lt $candidates.Count; $i++) {
        $ys = @($found[$candidates[$i]] | Sort-Object)
        Write-Host ("   [{0}] {1}   （{2}年度～{3}年度、{4}フォルダ）" -f ($i + 1), $candidates[$i], $ys[0], $ys[-1], $ys.Count)
    }
    Write-Host ''
    while (-not $target) {
        $ans = Read-Host '  成果物の場所の番号を入力してEnter（一覧にない場合はフォルダのパスを直接入力）'
        if ($ans -match '^\d+$' -and [int]$ans -ge 1 -and [int]$ans -le $candidates.Count) {
            $target = $candidates[[int]$ans - 1]
        } elseif ($ans -and (Test-Path -LiteralPath $ans.Trim('"'))) {
            $target = $ans.Trim('"')
        } else {
            Write-Host '  番号かパスが正しくありません。もう一度入力してください。' -ForegroundColor Yellow
        }
    }
} else {
    Write-Host '  年度フォルダが自動では見つかりませんでした。' -ForegroundColor Yellow
    while (-not $target) {
        $ans = Read-Host '  「2006年度」などのフォルダが入っている親フォルダのパスを入力してEnter'
        if ($ans -and (Test-Path -LiteralPath $ans.Trim('"'))) { $target = $ans.Trim('"') }
        else { Write-Host '  そのフォルダは見つかりません。' -ForegroundColor Yellow }
    }
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
    outputDir    = $out.Trim('"')
}
$config | ConvertTo-Json | Set-Content -Path (Join-Path $installDir 'config.json') -Encoding UTF8
Write-Host "  保存しました: $installDir\config.json"

# --- 4. 自動実行の登録と1回目の収集 ---
Write-Step 4 '自動実行を登録し、1回目の収集を行います'
& (Join-Path $installDir 'install.ps1')
& (Join-Path $installDir 'collect.ps1')

# 何の種類のファイルが集まったかを表示する（どのソフトの出力が含まれているかの確認用）
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
