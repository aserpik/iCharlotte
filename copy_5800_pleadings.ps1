$ErrorActionPreference = 'Continue'
$root = "Z:\Shared\Current Clients\5800 - AMTRUST"
$dest = "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs"
if (-not (Test-Path $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }

$manifest = Join-Path $dest "_manifest.csv"
"Matter,SourcePath,DestFile" | Out-File -FilePath $manifest -Encoding utf8

$copied = 0
$matters = Get-ChildItem $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '0 - ADMIN' }
$total = $matters.Count
$i = 0
foreach ($m in $matters) {
    $i++
    Write-Host "[$i/$total] $($m.Name)"
    # all pleading dirs in this matter
    $pleadDirs = Get-ChildItem $m.FullName -Recurse -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'plead' }
    if (-not $pleadDirs) { continue }
    # keep only top-level pleading dirs (not nested under another matched dir)
    $top = $pleadDirs | Where-Object {
        $d = $_
        -not ($pleadDirs | Where-Object { $_.FullName -ne $d.FullName -and $d.FullName.StartsWith($_.FullName + '\') })
    }
    foreach ($pd in $top) {
        $pdfs = Get-ChildItem $pd.FullName -Recurse -File -Filter *.pdf -ErrorAction SilentlyContinue
        foreach ($f in $pdfs) {
            $base = "$($m.Name)__$($f.Name)"
            $target = Join-Path $dest $base
            $n = 1
            while (Test-Path $target) {
                $stem = [System.IO.Path]::GetFileNameWithoutExtension($base)
                $ext  = [System.IO.Path]::GetExtension($base)
                $target = Join-Path $dest "$stem`_$n$ext"
                $n++
            }
            try {
                Copy-Item -LiteralPath $f.FullName -Destination $target -ErrorAction Stop
                $copied++
                $line = '"{0}","{1}","{2}"' -f $m.Name, $f.FullName, [System.IO.Path]::GetFileName($target)
                $line | Out-File -FilePath $manifest -Append -Encoding utf8
            } catch {
                Write-Host "  FAILED: $($f.FullName) -> $($_.Exception.Message)"
            }
        }
    }
}
Write-Host "=== DONE. Copied $copied PDFs to $dest ==="
