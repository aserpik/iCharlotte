$ErrorActionPreference = 'Continue'
$dest = "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs"
$manifest = Import-Csv (Join-Path $dest "_manifest.csv")
$ok = 0; $fail = 0; $skip = 0; $i = 0; $tot = $manifest.Count
foreach ($row in $manifest) {
    $i++
    $src = $row.SourcePath
    $target = Join-Path $dest $row.DestFile
    if (Test-Path -LiteralPath $target) { $skip++; continue }
    if (-not (Test-Path -LiteralPath $src)) { Write-Host "MISSING SRC: $src"; $fail++; continue }
    try {
        Copy-Item -LiteralPath $src -Destination $target -ErrorAction Stop
        $ok++
    } catch {
        Write-Host "FAILED: $src -> $($_.Exception.Message)"; $fail++
    }
    if ($i % 500 -eq 0) { Write-Host "[$i/$tot] copied=$ok skip=$skip fail=$fail" }
}
Write-Host "=== RECOPY DONE. copied=$ok skip=$skip fail=$fail  on-disk-pdfs=$((Get-ChildItem $dest -File -Filter *.pdf).Count) ==="
