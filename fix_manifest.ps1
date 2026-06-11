$dest = "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs"
$manifestPath = Join-Path $dest "_manifest.csv"
$manifest = Import-Csv $manifestPath

$dupNames = ($manifest | Group-Object DestFile | Where-Object { $_.Count -gt 1 }).Name
$keep     = $manifest | Where-Object { $dupNames -notcontains $_.DestFile }
$affected = $manifest | Where-Object { $dupNames -contains $_.DestFile }

# Reproduce the suffixing the copy used (same order, in-memory used-set)
$used = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$fixed = New-Object System.Collections.Generic.List[object]
foreach ($row in $affected) {
    $base = "$($row.Matter)__$([System.IO.Path]::GetFileName($row.SourcePath))"
    $cand = $base; $i = 1
    while ($used.Contains($cand)) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($base)
        $ext  = [System.IO.Path]::GetExtension($base)
        $cand = "{0}_{1}{2}" -f $stem, $i, $ext
        $i++
    }
    [void]$used.Add($cand)
    $fixed.Add([pscustomobject]@{ Matter=$row.Matter; SourcePath=$row.SourcePath; DestFile=$cand })
}

$final = New-Object System.Collections.Generic.List[object]
foreach ($r in $keep)  { $final.Add([pscustomobject]@{ Matter=$r.Matter; SourcePath=$r.SourcePath; DestFile=$r.DestFile }) }
foreach ($r in $fixed) { $final.Add($r) }

$final | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8

# Verify: every manifest DestFile must exist on disk; counts must match
$onDisk = [System.Collections.Generic.HashSet[string]]::new([string[]](Get-ChildItem $dest -File -Filter *.pdf | Select-Object -ExpandProperty Name),[System.StringComparer]::OrdinalIgnoreCase)
$missing = $final | Where-Object { -not $onDisk.Contains($_.DestFile) }
$dupCheck = $final | Group-Object DestFile | Where-Object { $_.Count -gt 1 }
"Manifest rows: $($final.Count)"
"PDFs on disk: $($onDisk.Count)"
"Manifest entries missing on disk: $($missing.Count)"
"Duplicate dest names remaining: $($dupCheck.Count)"
"Distinct source paths: $(($final | Select-Object -ExpandProperty SourcePath -Unique).Count)"
