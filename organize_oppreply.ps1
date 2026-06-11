param([switch]$Apply, [string]$Root = "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs")
$root = $Root
$targets = @('Oppositions','Replies')

# Ordered; first match wins. Matched against normalized filename. Ex Parte wins over the
# underlying relief (opposing/replying to an ex parte is its own form).
$rules = @(
    @('Non-Opposition Notices',     'non.?oppo|non.?opp\b|notice of non'),
    @('Ex Parte',                   'ex.?parte|\bEPA\b'),
    @('MSJ-MSA',                    '\bMSJ\b|\bMSA\b|summary judgment|summary adjudication|undisputed facts|separate stmt of undisputed|additional material facts|\bAMF\b'),
    @('Demurrer',                   'demurrer|\bdemur|\bdem\b'),
    @('Motion to Strike',           '\bMTS\b|motion to strike|to strike|strike punitive|strike complaint|strike portions'),
    @('Motion to Compel',           '\bMTC\b|\bMTCA\b|\bMTCF\b|motion to compel|to compel|deem admitted|motion to deem'),
    @('Motion in Limine',           'in limine|\bMIL\b'),
    @('Motion to Quash',            'quash'),
    @('Motion for Leave',           'leave to file|leave to amend|motion for leave|for leave'),
    @('Motion to Dismiss',          'motion to dismiss|to dismiss|\bMTD\b'),
    @('Set Aside Default',          'set aside|set-?aside|setaside'),
    @('Protective Order',           'protective order'),
    @('Continue Trial & Preference','continue trial|cont trial|cont\.? trial|trial continuance|trial preference|\bpreference\b|to continue|to cont\b')
)

foreach ($t in $targets) {
    $base = Join-Path $root $t
    $files = Get-ChildItem $base -File -Filter *.pdf
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($f in $files) {
        $orig = ($f.Name -split '__',2)[1]; if (-not $orig) { $orig = $f.Name }
        $norm = [System.IO.Path]::GetFileNameWithoutExtension($orig) -replace '_',' '
        $cat = 'Other'
        foreach ($r in $rules) { if ($norm -match $r[1]) { $cat = $r[0]; break } }
        $rows.Add([pscustomobject]@{ FileName=$f.Name; Sub=$cat })
    }
    "===== $t ($($rows.Count)) ====="
    $rows | Group-Object Sub | Sort-Object Count -Descending | ForEach-Object { "{0,4}  {1}" -f $_.Count, $_.Name }

    if ($Apply) {
        foreach ($row in $rows) {
            $folder = Join-Path $base $row.Sub
            if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
            if (-not (Get-Item -LiteralPath $folder).PSIsContainer) { throw "Not a directory: $folder" }
            Move-Item -LiteralPath (Join-Path $base $row.FileName) -Destination $folder -Force
        }
        "  APPLIED: $($rows.Count) files moved."
    }
}
