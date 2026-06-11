param([switch]$Apply, [string]$Root = "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs")
$base = Join-Path $Root "Ex Parte Applications"

# Ordered; first match wins. Matched against normalized filename (underscores->spaces, no ext).
$rules = @(
    # --- Not applications themselves: notices/letters & supporting papers (checked first) ---
    @('Notices & Letters', 'ex parte notice|notice of ex parte|\bNOR\b|notice of ruling|\bntc\b|\bletter\b|\bltr\b|via email|e-?mail'),
    @('Supporting Docs',   'joinder|memo of points|points and authorities|\bP&A\b|declaration|\bdec\b|\bdec\d|exparte dec|\bexhibit|\bexh\b|\bexh\.|proposed order|\bPO on\b|\bPO\b'),
    # --- Applications by relief sought ---
    @('Seal Documents',        'seal|under seal'),
    @('Enforce Settlement',    'enforce'),
    @('Leave & Specially Set', 'specially set|leave to amend|leave to file|supp(lemental)? brief|request for leave'),
    @('Advance Hearing',       'advance|\badv\b'),
    @('Continue Trial',        'continue|cont trial|cont\.? trial|\bcont\b'),
    @('Shorten Time (OST)',    'shorten time|shorten|\bOST\b|order shortening')
)

$files = Get-ChildItem $base -File -Filter *.pdf
$rows = New-Object System.Collections.Generic.List[object]
foreach ($f in $files) {
    $orig = ($f.Name -split '__',2)[1]; if (-not $orig) { $orig = $f.Name }
    $norm = [System.IO.Path]::GetFileNameWithoutExtension($orig) -replace '_',' '
    $cat = 'Other'
    foreach ($r in $rules) { if ($norm -match $r[1]) { $cat = $r[0]; break } }
    $rows.Add([pscustomobject]@{ FileName=$f.Name; Sub=$cat })
}
$rows | Group-Object Sub | Sort-Object Count -Descending | ForEach-Object { "{0,4}  {1}" -f $_.Count, $_.Name }
"---- TOTAL: $($rows.Count) ----"

if ($Apply) {
    foreach ($row in $rows) {
        $folder = Join-Path $base $row.Sub
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
        if (-not (Get-Item -LiteralPath $folder).PSIsContainer) { throw "Not a directory: $folder" }
        Move-Item -LiteralPath (Join-Path $base $row.FileName) -Destination $folder -Force
    }
    "APPLIED: moved $($rows.Count) ex parte files into sub-folders."
}
