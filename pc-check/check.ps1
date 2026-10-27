# 成果物の格納チェック（ALLION-01 で30分ごとに自動実行）。
#  ① 作業経過ログで「完了」の作業について、「成果物の場所」に実際にファイルがあるかを照合する
#  ② 2026年度の案件フォルダごとに、件数・容量・最終更新・今日の更新件数を集計する
# 結果は Notion「📦 成果物 格納チェック（自動更新）」に書き込む。人は開かず、Claude が読んで答える。
# Windows PowerShell 5.1（Windows 10 / 11 標準）で動作。
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
    # Notion へ書き込まず、結果を画面に出すだけにする
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$yearPattern = '^\s*(\d{4})\s*年度'
$notionVersion = '2025-09-03'
$apiBase = if ($env:PC_CHECK_NOTION_API) { $env:PC_CHECK_NOTION_API } else { 'https://api.notion.com/v1' }
$sep = [IO.Path]::DirectorySeparatorChar

# ---------- Notion ----------

function Get-NotionToken {
    $t = [Environment]::GetEnvironmentVariable('ALLION_NOTION_TOKEN', 'User')
    if (-not $t) { $t = $env:ALLION_NOTION_TOKEN }
    return $t
}

function Invoke-Notion($method, $path, $body) {
    $params = @{
        Uri         = "$apiBase$path"
        Method      = $method
        Headers     = @{ Authorization = "Bearer $script:token"; 'Notion-Version' = $notionVersion }
        ContentType = 'application/json; charset=utf-8'
    }
    if ($null -ne $body) {
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Depth 20 -Compress))
    }
    for ($try = 1; ; $try++) {
        try {
            $r = Invoke-RestMethod @params
            Start-Sleep -Milliseconds 350   # Notion の上限（平均 毎秒3回）を超えないように
            return $r
        } catch {
            $status = 0
            if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
            if (($status -eq 429 -or $status -ge 500) -and $try -lt 4) {
                Start-Sleep -Seconds ([math]::Pow(2, $try))
                continue
            }
            throw
        }
    }
}

function Get-AllRows($dataSourceId, $filter) {
    $rows = New-Object System.Collections.Generic.List[object]
    $cursor = $null
    do {
        $b = @{ page_size = 100 }
        if ($filter) { $b.filter = $filter }
        if ($cursor) { $b.start_cursor = $cursor }
        $r = Invoke-Notion 'POST' "/data_sources/$dataSourceId/query" $b
        foreach ($x in $r.results) { $rows.Add($x) }
        $cursor = $r.next_cursor
    } while ($r.has_more)
    return $rows
}

function Get-PlainText($prop) {
    if (-not $prop) { return '' }
    $parts = if ($prop.type -eq 'title') { $prop.title } else { $prop.rich_text }
    return (@($parts) | ForEach-Object { $_.plain_text }) -join ''
}

# 配列を返すときは「,」を付ける（付けないと PowerShell が1要素の配列をほどき、Notion が受け付けない形になる）
function New-Text($s) {
    if (-not $s) { return ,@() }
    if ($s.Length -gt 1900) { $s = $s.Substring(0, 1900) + '…' }
    return ,@(@{ text = @{ content = $s } })
}

function New-Date($dt) {
    if (-not $dt) { return $null }
    return @{ start = $dt.ToString('yyyy-MM-ddTHH:mm:sszzz') }
}

function Get-Md5($s) {
    $md5 = [Security.Cryptography.MD5]::Create()
    return -join ($md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($s)) | ForEach-Object { $_.ToString('x2') })
}

# ---------- パスの扱い ----------

# 割り当てドライブ（N: など）は自動実行のときに見えないことがあるため、\\サーバー\共有 に直す。
# 割り当てドライブでなければ $null（＝このPCのローカルドライブ）を返す
function ConvertTo-NetworkPath($p) {
    if ($p.StartsWith('\\')) { return $p }
    if ($p -match '^([A-Za-z]):\\?(.*)$') {
        $letter = $Matches[1].ToUpper(); $rest = $Matches[2]
        $remote = (Get-ItemProperty -LiteralPath "HKCU:\Network\$letter" -Name RemotePath -ErrorAction SilentlyContinue).RemotePath
        if ($remote) {
            # 「N:\Allion_Drive\01_社内共有\…」のように共有名まで書かれていることがある（N: が Allion_Drive そのもの）
            $share = $remote.TrimEnd('\').Split('\')[-1]
            if ($rest -like "$share\*") { $rest = $rest.Substring($share.Length + 1) }
            return $remote.TrimEnd('\') + '\' + $rest
        }
        # 再接続しない設定で割り当てたドライブは登録情報がないので、ドライブの種類で判断する
        $drive = New-Object IO.DriveInfo "${letter}:"
        if ($drive.DriveType -eq [IO.DriveType]::Network) { return $p }
    }
    return $null
}

# 「成果物の場所」の文章から、ドライブ文字または \\ で始まるパスを取り出す
function Get-PathsFromText($text) {
    $found = @()
    foreach ($m in [regex]::Matches($text, '(?<![A-Za-z0-9])(?:[A-Za-z]:\\|\\\\)[^\r\n、，,；;（「『／|]+')) {
        # 半角の ( ) はフォルダ名に使われる（例：飛行記録(ログ)）ので区切りにしない
        $found += $m.Value.Trim().TrimEnd('。', '.', ' ', '」')
    }
    return $found
}

# パスの後ろに「と 同名.zip」のような説明が続いていることがあるので、見つかるまで後ろの語を外して試す
function Resolve-ExistingPath($p, $testPath) {
    $cand = $p
    while ($cand) {
        if (& $testPath $cand) { return $cand }
        $i = $cand.LastIndexOf(' ')
        if ($i -lt 3) { break }
        $cand = $cand.Substring(0, $i).TrimEnd()
    }
    return $null
}

function Test-ExcludedName($name, $patterns) {
    foreach ($pattern in $patterns) {
        if ($name -like $pattern) { return $true }
    }
    return $false
}

function New-Stat {
    return @{ Count = 0; Bytes = [long]0; Last = $null; Today = 0 }
}

# フォルダの中を数える。作成日時・更新日時の新しい方が今日なら「今日の更新」に数える
function Add-Tree([IO.DirectoryInfo]$dir, $stat) {
    try { $items = $dir.GetFileSystemInfos() } catch { return }
    foreach ($i in $items) {
        if ($i -is [IO.DirectoryInfo]) {
            if ($i.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            if (Test-ExcludedName $i.Name $script:excludeDirs) { continue }
            Add-Tree $i $stat
        } else {
            if (Test-ExcludedName $i.Name $script:excludeNames) { continue }
            Add-File $i $stat
        }
    }
}

function Add-File([IO.FileInfo]$f, $stat) {
    $stat.Count++
    $stat.Bytes += $f.Length
    $t = if ($f.CreationTime -gt $f.LastWriteTime) { $f.CreationTime } else { $f.LastWriteTime }
    if (-not $stat.Last -or $t -gt $stat.Last) { $stat.Last = $t }
    if ($t -ge $script:today) { $stat.Today++ }
}

function Get-StatOfPath($p) {
    $stat = New-Stat
    $item = Get-Item -LiteralPath $p -Force
    if ($item -is [IO.DirectoryInfo]) { Add-Tree $item $stat } else { Add-File $item $stat }
    return $stat
}

# ---------- ① 作業経過ログの照合 ----------

function Test-LogRow($row) {
    $p = $row.properties
    $where = Get-PlainText $p.'成果物の場所'
    $paths = @(Get-PathsFromText $where)
    $result = @{
        Name = Get-PlainText $p.'作業名'; Machine = $p.'マシン'.select.name; Url = $row.url
        Where = $where; Judge = ''; Detail = ''; Stat = $null
    }
    if ($paths.Count -eq 0) {
        if ($where.Trim().StartsWith('なし')) {
            $result.Judge = '➖ 成果物なし'
            $result.Detail = '成果物の場所が「なし」の作業'
        } elseif ($where -match '(?<![A-Za-z0-9])[A-Za-z]:[^\\\s]') {
            $result.Judge = '⚠ 確認できない'
            $result.Detail = '成果物の場所のパスに「\」が抜けていて、場所を特定できない'
        } else {
            $result.Judge = '⚠ 確認できない'
            $result.Detail = '成果物の場所にフォルダのパスが書かれていない'
        }
        return $result
    }
    $stat = New-Stat
    $missing = @(); $local = @(); $unclear = @(); $ok = 0
    foreach ($path in $paths) {
        if ($path.Contains('…') -or $path.Contains('...')) { $unclear += $path; continue }
        if (-not (ConvertTo-NetworkPath $path)) { $local += $path; continue }
        $hit = Resolve-ExistingPath $path { param($x) Test-Path -LiteralPath (ConvertTo-NetworkPath $x) }
        if (-not $hit) { $missing += $path; continue }
        $s = Get-StatOfPath (ConvertTo-NetworkPath $hit)
        $stat.Count += $s.Count; $stat.Bytes += $s.Bytes; $stat.Today += $s.Today
        if ($s.Last -and (-not $stat.Last -or $s.Last -gt $stat.Last)) { $stat.Last = $s.Last }
        $ok++
    }
    $notes = @()
    if ($missing.Count) { $notes += '見つからない: ' + ($missing -join ' ／ ') }
    if ($unclear.Count) { $notes += '「…」で省略されていて確認できない: ' + ($unclear -join ' ／ ') }
    if ($local.Count) { $notes += '各PCのローカル（N:等の共有ではない）のため確認できない: ' + ($local -join ' ／ ') }
    if ($ok -gt 0 -and $stat.Count -eq 0) { $notes += 'フォルダはあるが中身が0件' }
    $result.Detail = $notes -join '。'
    $result.Stat = $stat
    if ($missing.Count -or ($ok -gt 0 -and $stat.Count -eq 0)) { $result.Judge = '❌ 見つからない' }
    elseif ($ok -eq 0) { $result.Judge = '⚠ 確認できない' }
    else { $result.Judge = '✅ 格納済み' }
    return $result
}

# ---------- ② 案件フォルダの集計 ----------

function Get-ProjectStats($root) {
    $stats = [ordered]@{}
    foreach ($year in Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue) {
        if ($year.Name -notmatch $yearPattern -or [int]$Matches[1] -lt $script:fromYear) { continue }
        foreach ($cat in @(Get-ChildItem -LiteralPath $year.FullName -Force -ErrorAction SilentlyContinue)) {
            if (-not ($cat -is [IO.DirectoryInfo])) {
                $key = "$($year.Name)$sep(年度直下のファイル)"
                if (-not $stats.Contains($key)) { $stats[$key] = New-Stat }
                if (-not (Test-ExcludedName $cat.Name $script:excludeNames)) { Add-File $cat $stats[$key] }
                continue
            }
            foreach ($proj in @(Get-ChildItem -LiteralPath $cat.FullName -Force -ErrorAction SilentlyContinue)) {
                if ($proj -is [IO.DirectoryInfo]) {
                    if (Test-ExcludedName $proj.Name $script:excludeDirs) { continue }
                    $key = "$($year.Name)$sep$($cat.Name)$sep$($proj.Name)"
                    $stats[$key] = New-Stat
                    Add-Tree $proj $stats[$key]
                } else {
                    $key = "$($year.Name)$sep$($cat.Name)$sep(直下のファイル)"
                    if (-not $stats.Contains($key)) { $stats[$key] = New-Stat }
                    if (-not (Test-ExcludedName $proj.Name $script:excludeNames)) { Add-File $proj $stats[$key] }
                }
            }
        }
    }
    return $stats
}

# ---------- 書き込み ----------

function Set-Row($key, $props, $existing) {
    $sig = Get-Md5 ($props | ConvertTo-Json -Depth 20 -Compress)
    $props['キー'] = @{ rich_text = (New-Text "$key|$sig") }
    $old = $existing[$key]
    if ($old -and $old.Sig -eq $sig) { return 'same' }
    if ($script:DryRun) { return 'dry' }
    if ($old) {
        Invoke-Notion 'PATCH' "/pages/$($old.Id)" @{ properties = $props } | Out-Null
        return 'updated'
    }
    Invoke-Notion 'POST' '/pages' @{ parent = @{ type = 'data_source_id'; data_source_id = $script:checkDs }; properties = $props } | Out-Null
    return 'created'
}

function New-Props($name, $kind, $judge, $machine, $where, $stat, $detail, $url) {
    $props = [ordered]@{
        '名前'   = @{ title = (New-Text $name) }
        '種類'   = @{ select = @{ name = $kind } }
        '判定'   = @{ select = @{ name = $judge } }
        '場所'   = @{ rich_text = (New-Text $where) }
        '詳細'   = @{ rich_text = (New-Text $detail) }
    }
    if ($machine) { $props['マシン'] = @{ select = @{ name = $machine } } }
    if ($url) { $props['作業経過ログ'] = @{ url = $url } }
    if ($stat) {
        $props['件数'] = @{ number = $stat.Count }
        $props['容量GB'] = @{ number = [math]::Round($stat.Bytes / 1GB, 2) }
        $props['今日の更新件数'] = @{ number = $stat.Today }
        $props['最終更新'] = @{ date = (New-Date $stat.Last) }
    }
    return $props
}

# ---------- 本体 ----------

function Invoke-Check {
    $started = Get-Date
    $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $script:fromYear = [int]$config.fromYear
    $script:excludeNames = @(@($config.excludeNames) | Where-Object { $_ })
    $script:excludeDirs = @(@($config.excludeDirs) | Where-Object { $_ })
    $script:checkDs = $config.checkDataSourceId
    $script:today = (Get-Date).Date
    $script:token = Get-NotionToken
    if (-not $script:token) { throw 'Notion の連携用の鍵（ALLION_NOTION_TOKEN）が設定されていません。setup.bat を実行してください。' }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $target = $config.target
    $display = if ($config.displayRoot) { $config.displayRoot } else { $target }

    # 前回の書き込み内容（キー｜内容の要約値）を読み、変わった行だけを書き直す
    $existing = @{}
    foreach ($row in Get-AllRows $script:checkDs $null) {
        $k = Get-PlainText $row.properties.'キー'
        if ($k -match '^(.*)\|([0-9a-f]{32})$') { $existing[$Matches[1]] = @{ Id = $row.id; Sig = $Matches[2] } }
    }
    $counts = @{}
    function Add-Count($what) { $counts[$what] = 1 + [int]$counts[$what] }

    # ① 作業経過ログ（直近 logDays 日に更新された「完了」の行）
    $since = (Get-Date).AddDays(-[int]$config.logDays).ToString('yyyy-MM-ddTHH:mm:sszzz')
    $filter = @{ and = @(
        @{ property = '状態'; select = @{ equals = '完了' } },
        @{ timestamp = 'last_edited_time'; last_edited_time = @{ on_or_after = $since } }
    ) }
    $logRows = Get-AllRows $config.logDataSourceId $filter
    $summary = @{ '✅ 格納済み' = 0; '❌ 見つからない' = 0; '⚠ 確認できない' = 0; '➖ 成果物なし' = 0 }
    foreach ($row in $logRows) {
        $r = Test-LogRow $row
        $summary[$r.Judge]++
        $props = New-Props $r.Name '作業ログ照合' $r.Judge $r.Machine $r.Where $r.Stat $r.Detail $r.Url
        Add-Count (Set-Row "LOG:$($row.id)" $props $existing)
        if ($script:DryRun) { Write-Output ("[{0}] {1}  {2}" -f $r.Judge, $r.Name, $r.Detail) }
    }

    # ② 案件フォルダ
    $stats = Get-ProjectStats $target
    $todayProjects = 0
    foreach ($key in $stats.Keys) {
        $s = $stats[$key]
        $judge = if ($s.Today -gt 0) { '📁 今日更新あり' } else { '📁 変化なし' }
        if ($s.Today -gt 0) { $todayProjects++ }
        $name = ($key.Split($sep) | Select-Object -Skip 1) -join '／'
        $props = New-Props $name '案件フォルダ' $judge $null ("$display$sep$key") $s '' $null
        Add-Count (Set-Row "DIR:$key" $props $existing)
        if ($script:DryRun -and $s.Today -gt 0) { Write-Output ("[{0}] {1}  今日 {2}件" -f $judge, $name, $s.Today) }
    }

    # 実行記録（毎回更新。Claude はここで「いつ時点の情報か」を確かめる）
    $secs = [int]((Get-Date) - $started).TotalSeconds
    $detail = "作業ログ照合 $($logRows.Count)件（✅$($summary['✅ 格納済み']) ❌$($summary['❌ 見つからない']) ⚠$($summary['⚠ 確認できない']) ➖$($summary['➖ 成果物なし'])）／" +
        "案件フォルダ $($stats.Count)件（今日更新あり ${todayProjects}件）／所要 ${secs}秒／実行PC $env:COMPUTERNAME"
    $run = New-Props '⏱ 最終チェック' 'チェック実行' '⏱ 実行記録' $null $display $null $detail $null
    $run['最終更新'] = @{ date = (New-Date (Get-Date)) }
    Add-Count (Set-Row 'RUN' $run $existing)

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $detail（Notion 新規$([int]$counts['created'])・更新$([int]$counts['updated'])・変化なし$([int]$counts['same'])）"
    return $line
}

# テストで関数だけ読み込むとき（. .\check.ps1）は本体を動かさない
if ($MyInvocation.InvocationName -ne '.') {
    $logFile = Join-Path $PSScriptRoot 'check.log'
    try {
        $out = @(Invoke-Check)
        $out | ForEach-Object { Write-Output $_ }
        if (-not $DryRun) { Set-Content -LiteralPath $logFile -Value $out[-1] -Encoding UTF8 }
    } catch {
        $msg = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') エラー: $($_.Exception.Message)"
        Set-Content -LiteralPath $logFile -Value $msg -Encoding UTF8
        Write-Error $msg
        exit 1
    }
}
