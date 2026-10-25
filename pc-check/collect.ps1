# 各PCの成果物ファイル一覧を収集し、共有フォルダに「PC名.csv」として書き出す。
# Windows PowerShell 5.1（Windows 10 / 11 標準）で動作。
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json')
)

$ErrorActionPreference = 'Stop'

$config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$pc = $env:COMPUTERNAME
$user = $env:USERNAME
$collectedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$extensions = @($config.extensions | ForEach-Object { $_.ToLower() })
$excludeNames = @($config.excludeNames)
$outputDir = [Environment]::ExpandEnvironmentVariables($config.outputDir)

$rows = New-Object System.Collections.Generic.List[object]

foreach ($rawTarget in $config.targets) {
    $target = [Environment]::ExpandEnvironmentVariables($rawTarget)
    if (-not (Test-Path -LiteralPath $target)) {
        Write-Warning "対象フォルダが見つかりません: $target"
        continue
    }
    $root = (Resolve-Path -LiteralPath $target).Path.TrimEnd('\')

    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Force -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        if ($extensions.Count -gt 0 -and $extensions -notcontains $f.Extension.ToLower()) { continue }
        $skip = $false
        foreach ($pattern in $excludeNames) {
            if ($f.Name -like $pattern) { $skip = $true; break }
        }
        if ($skip) { continue }

        $rows.Add([pscustomobject]@{
            PC           = $pc
            User         = $user
            Root         = $root
            RelativePath = $f.FullName.Substring($root.Length).TrimStart('\')
            Name         = $f.Name
            Extension    = $f.Extension.ToLower()
            SizeKB       = [math]::Round($f.Length / 1KB, 1)
            Modified     = $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            Created      = $f.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
            FullPath     = $f.FullName
            CollectedAt  = $collectedAt
        })
    }
}

if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# 書き出し途中のファイルをビューアが読まないよう、一時ファイルに書いてから置き換える
$outFile = Join-Path $outputDir "$pc.csv"
$tmpFile = "$outFile.tmp"
$header = 'PC,User,Root,RelativePath,Name,Extension,SizeKB,Modified,Created,FullPath,CollectedAt'
if ($rows.Count -gt 0) {
    $rows | Export-Csv -Path $tmpFile -NoTypeInformation -Encoding UTF8
} else {
    # 該当ファイル0件でも「収集は動いている」ことが分かるよう、ヘッダーと収集時刻だけの行を残す
    $lines = @($header, "`"$pc`",`"$user`",,,,,,,,,`"$collectedAt`"")
    Set-Content -Path $tmpFile -Value $lines -Encoding UTF8
}
Move-Item -LiteralPath $tmpFile -Destination $outFile -Force

Write-Output "$($rows.Count) 件を書き出しました: $outFile"
