# 各PCの成果物ファイル一覧を収集し、出力先フォルダに「PC名.csv」として書き出す。
# 成果物は「N:\01_社内共有\01_案件データ\2026年度\ドローン 等\案件フォルダ\...」の形で格納されている前提で、
# fromYear 以降の年度フォルダの中だけを対象にし、年度・年度直下のフォルダ・案件フォルダも列に出す。
# Windows PowerShell 5.1（Windows 10 / 11 標準）で動作。
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json')
)

$ErrorActionPreference = 'Stop'

$config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$pc = $env:COMPUTERNAME
$user = $env:USERNAME
$collectedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$extensions = @(@($config.extensions) | Where-Object { $_ } | ForEach-Object { $_.ToLower() })
$excludeNames = @(@($config.excludeNames) | Where-Object { $_ })
$excludeDirs = @(@($config.excludeDirs) | Where-Object { $_ })
$fromYear = if ($config.fromYear) { [int]$config.fromYear } else { 0 }
$outputDir = [Environment]::ExpandEnvironmentVariables($config.outputDir)
$yearPattern = '^\s*(\d{4})\s*年度'
# 年度フォルダを探しにいく深さ（対象フォルダの直下に年度フォルダがあれば 0 で足りる。余裕を見て 3）
$searchDepth = 3

$rows = New-Object System.Collections.Generic.List[object]

function Add-File($f, $root) {
    if ($extensions.Count -gt 0 -and $extensions -notcontains $f.Extension.ToLower()) { return }
    foreach ($pattern in $excludeNames) {
        if ($f.Name -like $pattern) { return }
    }
    $relative = $f.FullName.Substring($root.Length).TrimStart('\')
    $parts = $relative.Split('\')
    # パスの中で最初に出てくる「NNNN年度」フォルダを年度とし、その直下を Category、さらにその下を Project とする
    $year = ''; $category = ''; $project = ''
    for ($i = 0; $i -lt $parts.Length - 1; $i++) {
        if ($parts[$i] -match $yearPattern) {
            $year = $Matches[1]
            if ($i + 1 -lt $parts.Length - 1) { $category = $parts[$i + 1] }
            if ($i + 2 -lt $parts.Length - 1) { $project = $parts[$i + 2] }
            break
        }
    }
    $rows.Add([pscustomobject]@{
        PC           = $pc
        User         = $user
        Root         = $root
        RelativePath = $relative
        Name         = $f.Name
        Extension    = $f.Extension.ToLower()
        SizeKB       = [math]::Round($f.Length / 1KB, 1)
        Modified     = $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
        Created      = $f.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
        FullPath     = $f.FullName
        Year         = $year
        Category     = $category
        Project      = $project
        CollectedAt  = $collectedAt
    })
}

# 年度フォルダの外は年度フォルダを探すだけ（深さ制限つき）、年度フォルダの中は全ファイルを集める
function Find-Files($dir, $root, $depth, $inYear) {
    foreach ($item in Get-ChildItem -LiteralPath $dir -Force -ErrorAction SilentlyContinue) {
        if (-not $item.PSIsContainer) {
            if ($inYear) { Add-File $item $root }
            continue
        }
        # Metashape の *.files など、ソフト内部の作業フォルダは中を見ない
        $skip = $false
        foreach ($pattern in $excludeDirs) {
            if ($item.Name -like $pattern) { $skip = $true; break }
        }
        if ($skip) { continue }
        if ($inYear) {
            Find-Files $item.FullName $root ($depth + 1) $true
        } elseif ($item.Name -match $yearPattern) {
            if ([int]$Matches[1] -ge $fromYear) { Find-Files $item.FullName $root ($depth + 1) $true }
        } elseif ($depth -lt $searchDepth) {
            Find-Files $item.FullName $root ($depth + 1) $false
        }
    }
}

foreach ($rawTarget in $config.targets) {
    $target = [Environment]::ExpandEnvironmentVariables($rawTarget)
    if (-not (Test-Path -LiteralPath $target)) {
        Write-Warning "対象フォルダが見つかりません: $target"
        continue
    }
    $root = (Resolve-Path -LiteralPath $target).ProviderPath.TrimEnd('\')
    Find-Files $root $root 0 $false
}

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 書き出し途中のファイルをビューアが読まないよう、一時ファイルに書いてから置き換える
$outFile = Join-Path $outputDir "$pc.csv"
$tmpFile = "$outFile.tmp"
$header = 'PC,User,Root,RelativePath,Name,Extension,SizeKB,Modified,Created,FullPath,Year,Category,Project,CollectedAt'
if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $tmpFile -NoTypeInformation -Encoding UTF8
} else {
    # 該当ファイル0件でも「収集は動いている」ことが分かるよう、ヘッダーと収集時刻だけの行を残す
    $lines = @($header, "`"$pc`",`"$user`",,,,,,,,,,,,`"$collectedAt`"")
    Set-Content -Path $tmpFile -Value $lines -Encoding UTF8
}
Move-Item -LiteralPath $tmpFile -Destination $outFile -Force

Write-Output "$($rows.Count) 件を書き出しました: $outFile"
