$ErrorActionPreference = 'Stop'
$dest = "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs"
$manifestPath = Join-Path $dest "_manifest.csv"
$manifest = Import-Csv $manifestPath

$dupNames = ($manifest | Group-Object DestFile | Where-Object { $_.Count -gt 1 }).Name
$affected = $manifest | Where-Object { $dupNames -contains $_.DestFile }
$keep     = $manifest | Where-Object { $dupNames -notcontains $_.DestFile }

Write-Host "Affected rows: $($affected.Count); dup dest names: $($dupNames.Count)"

# 1) Delete the collapsed dest files (literal)
foreach ($n in $dupNames) {
    $p = Join-Path $dest $n
    if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
}

# 2) Re-copy each affected source with LiteralPath collision handling
$newRows = New-Object System.Collections.Generic.List[object]
foreach ($row in $affected) {
    $src = $row.SourcePath
    $matter = $row.Matter
    $base = "$matter`__$([System.IO.Path]::GetFileName($src))"
    $target = Join-Path $dest $base
    $i = 1
    while (Test-Path -LiteralPath $target) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($base)
        $ext  = [System.IO.Path]::GetExtension($base)
        $target = Join-Path $dest ("{0}_{1}{2}" -f $stem, $i, $ext)
        $i++
    }
    if (-not (Test-Path -LiteralPath $src)) { Write-Host "MISSING SOURCE: $src"; continue }
    Copy-Item -LiteralPath $src -Destination $target -Force
    $newRows.Add([pscustomobject]@{ Matter=$matter; SourcePath=$src; DestFile=[System.IO.Path]::GetFileName($target) })
}

# 3) Rewrite manifest = kept rows + corrected rows
$final = @($keep) + @($newRows)
$final | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8
Write-Host "Recovered $($newRows.Count) files. New manifest rows: $($final.Count)"
