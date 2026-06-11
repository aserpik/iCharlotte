param([switch]$Apply, [string]$Dest = "C:\geminiterminal2\5800_AMTRUST_Pleadings_PDFs")

$dest = $Dest

# Ordered rules: first match wins. Patterns are case-insensitive regex matched against a
# NORMALIZED filename (extension dropped, underscores -> spaces so \b boundaries work).
$rules = @(
    # --- High-priority support docs (win over the motion they support) ---
    @('_Support - Declarations',        '\bdecl\b|declaration|\bdec\.? of\b|decl\. of'),
    @('_Support - Separate Statements', 'separate statement|\bSSUMF\b|\bSSAMF\b|\bSS ISO|statement of undisputed|sep\.? stmt'),
    @('_Support - Proposed Orders',     'proposed.{0,4}order|\[propos|\[prop\b|\[pp\b|\bpp order|\[po\]'),
    @('_Support - Proof of Service',    'proof of service|\bPOS\b'),
    @('_Support - Orders & Rulings',    'tentative ruling|\btentative\b|minute order|notice of ruling|\bruling\b|order granting|order denying|order after hearing|signed order|order re |order on |order entered|order shortening|\bOSC\b'),
    @('_Support - Evidentiary Objections','objections? to (the )?evidence|obj\.? to (the )?evidence|evidentiary objection'),
    @('_Support - Hearing Reservations','hearing reservation|\breservation\b|\bMJOP\b'),
    @('_Support - Notices',             'notice of (ruling|entry|hearing|continu|posting|lodging|errata|rejection|appearance|depo|change|order\b|withdrawal|non.?appear|reassign|related|settlement|stay|case|status|trial)'),
    # --- Amendments (before Complaint, so Doe amendments aren't filed as complaints) ---
    @('_Support - Amendments',          'doe amendment|amendment to (the )?(complaint|cc|cross|answer|ap)|\bdoe \d|fictitious'),
    # --- Briefs (Reply before Opposition; both before the motion they address) ---
    @('Replies',                        '\breply\b|reply in support|reply brief|reply memo'),
    @('Oppositions',                    '\boppos|\boppo|\bopp\b|opp\.? to\b|opp''?n|statement in opposition'),
    # --- Ex parte (procedurally its own form, regardless of underlying relief) ---
    @('Ex Parte Applications',          'ex.?parte'),
    # --- Substantive motions / moving papers ---
    @('Motion - Summary Judgment',      'summary judgment|summary adjudication|\bMSJ\b|\bMSA\b'),
    @('Motion - Demurrer',              'demurrer|\bdemur'),
    @('Motion - Strike',                'motion to strike|mtn to strike|\bMTS\b|strike portions|strike the|strike answer'),
    @('Motion - Compel',                'motion to compel|\bMTC\b|to compel'),
    @('Motion - In Limine',             'in limine|\bMIL\b'),
    @('Motion - Quash',                 'quash'),
    @('Motion - Sanctions',             'sanction'),
    @('Motion - Relieve Counsel',       'relieved as counsel|be relieved|withdraw as counsel|motion to withdraw'),
    @('Motion - Continue Trial',        'continue trial|cont.{0,8}trial|trial continuance|continuance of trial|advance.{0,10}trial|specially set|preferential|preference'),
    # --- Pleadings (Answer first so "Answer to Complaint" isn't filed as Complaint) ---
    @('Pleadings - Answer',             '\banswer\b'),
    @('Pleadings - Amended Complaint',  'amended complaint|first amended|second amended|third amended|\bFAC\b|\bSAC\b|\bTAC\b'),
    @('Pleadings - Cross-Complaint',    'cross.?complaint|\bXC\b|\bXCOMP|cross.?claim'),
    @('Pleadings - Complaint',          '\bcomplaint\b'),
    # --- Case initiation / default ---
    @('_Support - Default',             'entry of default|request for entry of default|default judgment|\bdefault\b'),
    @('_Support - Summons & Case Init', '\bsummons\b|civil case cover|case cover sheet|\bS&C\b|cover sheet'),
    # --- Discovery (before generic motion catch-all so "...Pleadings, Motions and Discovery" isn't grabbed as a motion) ---
    @('_Support - Discovery Responses', 'response to|\bROGS?\b|\bSROG|\bFROG|\bRFP\b|\bRFP\d|\bRFAs?\b|\bRPD\b|interrogator|request for production|request for admission|request for prior pleading|prior pleadings|further response|supplemental response|verification|doc(ument)? produced|expert (demand|disclosure|designation|witness)|demand for exchange'),
    # --- Generic motion catch-all (typed motions already handled above) ---
    @('Motions - Other',                '\bmotion\b|\bmtn\b|notice of motion|\bpetition\b|leave to amend'),
    # --- Remaining support / ancillary ---
    @('_Support - Substitution of Atty','substitution of att|sub of att|subofatt|subst.{0,5}att|\bSOA\b'),
    @('_Support - Statement of Damages','statement of damages'),
    @('_Support - Settlement & 998',    '\b998\b|offer to compromise|settlement agreement|settlement conference|settlement offer|settlement demand|\bmediation\b|\bMSC\b|mandatory settlement'),
    @('_Support - Jury Demand',         'jury trial|demand for jury|jury fee|jury demand'),
    @('_Support - Stipulations',        'stipulation|\bstip\b'),
    @('_Support - Case Management',     'case management|\bCMC\b|\bCMS\b|status conference|\bTSC\b|trial setting|case mgt|cm statement'),
    @('_Support - Request for Dismissal','dismissal|request for dismiss'),
    @('_Support - Subpoenas',           'subpoena|subpena'),
    @('_Support - Meet and Confer',     'meet and confer|\bm&c\b'),
    @('_Support - Exhibits',            '\bexhibit|\bexh\b|\bexh\.|\bex \d|\bexs?\b \d|compendium|exhibit packet|index of exhibit'),
    @('_Support - Dockets',             '\bdocket\b'),
    @('_Support - Amendments',          'doe amendment|amendment to (the )?(complaint|cc|cross|answer|ap)|\bamendment to\b'),
    @('_Support - Correspondence',      '\bletter\b|\bemail\b|\be-mail\b|correspondence'),
    @('_Support - Filing & Receipts',   'efile conf|e-?file|filing confirmation|\breceipt\b|\bNEF\b|electronic filing'),
    @('_Support - Notices',             '\bnotice\b|\bntc\b|\bnoc\b|\bnor\b|\bNOD\b|\bNOA\b|\bNOH\b|\bNRUL|\bNAR\b|noticeof|ntcof')
)

$files = Get-ChildItem $dest -File -Filter *.pdf
$rows = New-Object System.Collections.Generic.List[object]
foreach ($f in $files) {
    $orig = ($f.Name -split '__',2)[1]; if (-not $orig) { $orig = $f.Name }
    # normalize: drop extension, underscores -> spaces
    $norm = [System.IO.Path]::GetFileNameWithoutExtension($orig) -replace '_',' '
    $cat = '_Other'
    foreach ($r in $rules) { if ($norm -match $r[1]) { $cat = $r[0]; break } }
    $rows.Add([pscustomobject]@{ FileName=$f.Name; Category=$cat })
}

$rows | Group-Object Category | Sort-Object Count -Descending | ForEach-Object { "{0,5}  {1}" -f $_.Count, $_.Name }
"---- TOTAL: $($rows.Count) ----"
$rows | Export-Csv (Join-Path $dest "_classification_preview.csv") -NoTypeInformation -Encoding utf8

if ($Apply) {
    foreach ($row in $rows) {
        $folder = Join-Path $dest $row.Category
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }
        # Hard guard: never move unless the destination is a real directory (prevents file-rename collapse)
        if (-not (Get-Item -LiteralPath $folder).PSIsContainer) { throw "Destination is not a directory: $folder" }
        Move-Item -LiteralPath (Join-Path $dest $row.FileName) -Destination $folder -Force
    }
    "APPLIED: moved $($rows.Count) files into $(( $rows | Group-Object Category).Count) folders."
}
