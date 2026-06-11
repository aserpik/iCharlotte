$ErrorActionPreference = 'Continue'
$root = "Z:\Shared\Current Clients\3800- NATIONWIDE"
$dest = "C:\geminiterminal2\3800_NATIONWIDE_Pleadings_PDFs"
if (-not (Test-Path -LiteralPath $dest)) { New-Item -ItemType Directory -Path $dest | Out-Null }
$manifestPath = Join-Path $dest "_manifest.csv"
$MIN_FREE_GB = 2.0

# Classification rules (same as organize_pleadings.ps1). We KEEP only substantive categories
# (those NOT starting with "_") -> motions, oppositions, replies, ex parte, pleadings.
$rules = @(
    @('_Support - Declarations','\bdecl\b|declaration|\bdec\.? of\b|decl\. of'),
    @('_Support - Separate Statements','separate statement|\bSSUMF\b|\bSSAMF\b|\bSS ISO|statement of undisputed|sep\.? stmt'),
    @('_Support - Proposed Orders','proposed.{0,4}order|\[propos|\[prop\b|\[pp\b|\bpp order|\[po\]'),
    @('_Support - Proof of Service','proof of service|\bPOS\b'),
    @('_Support - Orders & Rulings','tentative ruling|\btentative\b|minute order|notice of ruling|\bruling\b|order granting|order denying|order after hearing|signed order|order re |order on |order entered|order shortening|\bOSC\b'),
    @('_Support - Evidentiary Objections','objections? to (the )?evidence|obj\.? to (the )?evidence|evidentiary objection'),
    @('_Support - Hearing Reservations','hearing reservation|\breservation\b|\bMJOP\b'),
    @('_Support - Notices','notice of (ruling|entry|hearing|continu|posting|lodging|errata|rejection|appearance|depo|change|order\b|withdrawal|non.?appear|reassign|related|settlement|stay|case|status|trial)'),
    @('_Support - Amendments','doe amendment|amendment to (the )?(complaint|cc|cross|answer|ap)|\bdoe \d|fictitious'),
    @('Replies','\breply\b|reply in support|reply brief|reply memo'),
    @('Oppositions','\boppos|\boppo|\bopp\b|opp\.? to\b|opp''?n|statement in opposition'),
    @('Ex Parte Applications','ex.?parte'),
    @('Motion - Summary Judgment','summary judgment|summary adjudication|\bMSJ\b|\bMSA\b'),
    @('Motion - Demurrer','demurrer|\bdemur'),
    @('Motion - Strike','motion to strike|mtn to strike|\bMTS\b|strike portions|strike the|strike answer'),
    @('Motion - Compel','motion to compel|\bMTC\b|to compel'),
    @('Motion - In Limine','in limine|\bMIL\b'),
    @('Motion - Quash','quash'),
    @('Motion - Sanctions','sanction'),
    @('Motion - Relieve Counsel','relieved as counsel|be relieved|withdraw as counsel|motion to withdraw'),
    @('Motion - Continue Trial','continue trial|cont.{0,8}trial|trial continuance|continuance of trial|advance.{0,10}trial|specially set|preferential|preference'),
    @('Pleadings - Answer','\banswer\b'),
    @('Pleadings - Amended Complaint','amended complaint|first amended|second amended|third amended|\bFAC\b|\bSAC\b|\bTAC\b'),
    @('Pleadings - Cross-Complaint','cross.?complaint|\bXC\b|\bXCOMP|cross.?claim'),
    @('Pleadings - Complaint','\bcomplaint\b'),
    @('_Support - Default','entry of default|request for entry of default|default judgment|\bdefault\b'),
    @('_Support - Summons & Case Init','\bsummons\b|civil case cover|case cover sheet|\bS&C\b|cover sheet'),
    @('_Support - Discovery Responses','response to|\bROGS?\b|\bSROG|\bFROG|\bRFP\b|\bRFP\d|\bRFAs?\b|\bRPD\b|interrogator|request for production|request for admission|request for prior pleading|prior pleadings|further response|supplemental response|verification|doc(ument)? produced|expert (demand|disclosure|designation|witness)|demand for exchange'),
    @('Motions - Other','\bmotion\b|\bmtn\b|notice of motion|\bpetition\b|leave to amend'),
    @('_Support - Substitution of Atty','substitution of att|sub of att|subofatt|subst.{0,5}att|\bSOA\b'),
    @('_Support - Statement of Damages','statement of damages'),
    @('_Support - Settlement & 998','\b998\b|offer to compromise|settlement agreement|settlement conference|settlement offer|settlement demand|\bmediation\b|\bMSC\b|mandatory settlement'),
    @('_Support - Jury Demand','jury trial|demand for jury|jury fee|jury demand'),
    @('_Support - Stipulations','stipulation|\bstip\b'),
    @('_Support - Case Management','case management|\bCMC\b|\bCMS\b|status conference|\bTSC\b|trial setting|case mgt|cm statement'),
    @('_Support - Request for Dismissal','dismissal|request for dismiss'),
    @('_Support - Subpoenas','subpoena|subpena'),
    @('_Support - Meet and Confer','meet and confer|\bm&c\b'),
    @('_Support - Exhibits','\bexhibit|\bexh\b|\bexh\.|\bex \d|\bexs?\b \d|compendium|exhibit packet|index of exhibit'),
    @('_Support - Dockets','\bdocket\b'),
    @('_Support - Correspondence','\bletter\b|\bemail\b|\be-mail\b|correspondence'),
    @('_Support - Filing & Receipts','efile conf|e-?file|filing confirmation|\breceipt\b|\bNEF\b|electronic filing'),
    @('_Support - Notices2','\bnotice\b|\bntc\b|\bnoc\b|\bnor\b|\bNOD\b|\bNOA\b|\bNOH\b|\bNRUL|\bNAR\b|noticeof|ntcof')
)
function Classify($name) {
    $norm = [System.IO.Path]::GetFileNameWithoutExtension($name) -replace '_',' '
    foreach ($r in $rules) { if ($norm -match $r[1]) { return $r[0] } }
    return '_Other'
}

# Phase A: enumerate, classify, KEEP substantive only, assign deterministic dest names
$matters = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '0 - ADMIN' }
$plan = New-Object System.Collections.Generic.List[object]
$used = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$mi=0; $mtot=$matters.Count; $seen=0; $kept=0
foreach ($m in $matters) {
    $mi++
    Write-Host "[enum $mi/$mtot] $($m.Name)"
    $pleadDirs = Get-ChildItem -LiteralPath $m.FullName -Recurse -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'plead' }
    if (-not $pleadDirs) { continue }
    $top = $pleadDirs | Where-Object { $d=$_; -not ($pleadDirs | Where-Object { $_.FullName -ne $d.FullName -and $d.FullName.StartsWith($_.FullName + '\') }) }
    $pdfs = foreach ($pd in $top) { Get-ChildItem -LiteralPath $pd.FullName -Recurse -File -Filter *.pdf -ErrorAction SilentlyContinue }
    foreach ($f in ($pdfs | Sort-Object FullName)) {
        $seen++
        $cat = Classify $f.Name
        if ($cat -match '^_') { continue }   # skip support/other
        $kept++
        $base = "$($m.Name)__$($f.Name)"
        $cand = $base; $n=1
        while ($used.Contains($cand)) {
            $stem=[System.IO.Path]::GetFileNameWithoutExtension($base); $ext=[System.IO.Path]::GetExtension($base)
            $cand = "{0}_{1}{2}" -f $stem,$n,$ext; $n++
        }
        [void]$used.Add($cand)
        $plan.Add([pscustomobject]@{ Matter=$m.Name; SourcePath=$f.FullName; DestFile=$cand; Category=$cat })
    }
}
Write-Host "ENUM DONE: seen=$seen kept(substantive)=$kept"
$plan | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding utf8

# Phase B: copy substantive files (skip-existing + disk guard)
$ok=0; $skip=0; $fail=0; $i=0; $tot=$plan.Count
foreach ($row in $plan) {
    $i++
    $target = Join-Path $dest $row.DestFile
    if (Test-Path -LiteralPath $target) { $skip++; continue }
    $free = (Get-PSDrive C).Free/1GB
    if ($free -lt $MIN_FREE_GB) { Write-Host "ABORT: C: free ${free}GB < ${MIN_FREE_GB}GB at [$i/$tot]. copied=$ok"; break }
    if (-not (Test-Path -LiteralPath $row.SourcePath)) { Write-Host "MISSING SRC: $($row.SourcePath)"; $fail++; continue }
    try { Copy-Item -LiteralPath $row.SourcePath -Destination $target -ErrorAction Stop; $ok++ }
    catch { Write-Host "FAILED: $($row.SourcePath) -> $($_.Exception.Message)"; $fail++ }
    if ($i % 500 -eq 0) { Write-Host "[copy $i/$tot] copied=$ok skip=$skip fail=$fail free=$([math]::Round((Get-PSDrive C).Free/1GB,1))GB" }
}
$onDisk = (Get-ChildItem -LiteralPath $dest -File -Filter *.pdf).Count
Write-Host "=== COPY DONE. planned=$tot copied=$ok skip=$skip fail=$fail on-disk=$onDisk free=$([math]::Round((Get-PSDrive C).Free/1GB,1))GB ==="
