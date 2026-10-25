# 各PCの成果物ファイル一覧を収集し、出力先フォルダに「PC名.csv」として書き出す。
# 成果物は「対象フォルダ\2006年度\分類フォルダ\...」の形で整理されている前提で、
# 年度と分類（年度直下のフォルダ名）も列に出す。
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
$fromYear = if ($config.fromYear) { [int]$config.fromYear } else { 0 }
$outputDir = [Environment]::ExpandEnvironmentVariables($config.outputDir)
$yearPattern = '^\s*(\d{4})\s*年度'

$rows = New-Object System.Collections.Generic.List[object]

function Add-File($f, $root) {
    if ($extensions.Count -gt 0 -and $extensions -notcontains $f.Extension.ToLower()) { return }
    foreach ($pattern in $excludeNames) {
        if ($f.Name -like $pattern) { return }
    }
    $relative = $f.FullName.Substring($root.Length).TrimStart('\')
    $parts = $relative.Split('\')
    # パスの中で最初に出てくる「NNNN年度」フォルダを年度とし、その直下のフォルダを分類とする
    $year = ''; $category = ''
    for ($i = 0; $i -lt $parts.Length - 1; $i++) {
        if ($parts[$i] -match $yearPattern) {
            $year = $Matches[1]
            if ($i + 1 -lt $parts.Length - 1) { $category = $parts[$i + 1] }
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
        CollectedAt  = $collectedAt
    })
}

foreach ($rawTarget in $config.targets) {
    $target = [Environment]::ExpandEnvironmentVariables($rawTarget)
    if (-not (Test-Path -LiteralPath $target)) {
        Write-Warning "対象フォルダが見つかりません: $target"
        continue
    }
    $root = (Resolve-Path -LiteralPath $target).Path.TrimEnd('\')

    foreach ($item in Get-ChildItem -LiteralPath $root -Force -ErrorAction SilentlyContinue) {
        if ($item.PSIsContainer) {
            # fromYear より古い年度フォルダは中を見ない
            if ($fromYear -and $item.Name -match $yearPattern -and [int]$Matches[1] -lt $fromYear) { continue }
            foreach ($f in Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue) {
                Add-File $f $root
            }
        } else {
            Add-File $item $root
        }
    }
}

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 書き出し途中のファイルをビューアが読まないよう、一時ファイルに書いてから置き換える
$outFile = Join-Path $outputDir "$pc.csv"
$tmpFile = "$outFile.tmp"
$header = 'PC,User,Root,RelativePath,Name,Extension,SizeKB,Modified,Created,FullPath,Year,Category,CollectedAt'
if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $tmpFile -NoTypeInformation -Encoding UTF8
} else {
    # 該当ファイル0件でも「収集は動いている」ことが分かるよう、ヘッダーと収集時刻だけの行を残す
    $lines = @($header, "`"$pc`",`"$user`",,,,,,,,,,,`"$collectedAt`"")
    Set-Content -Path $tmpFile -Value $lines -Encoding UTF8
}
Move-Item -LiteralPath $tmpFile -Destination $outFile -Force

Write-Output "$($rows.Count) 件を書き出しました: $outFile"
