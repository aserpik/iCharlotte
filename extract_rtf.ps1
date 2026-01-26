Add-Type -AssemblyName System.Windows.Forms
Write-Host "Assembly loaded."
$rtf = New-Object System.Windows.Forms.RichTextBox
$path = "C:\geminiterminal2\LLM Resources\Calendaring\LLM Scan"
$files = Get-ChildItem $path -Filter "*.doc"
$output = ""

foreach ($file in $files) {
    if ($file.Name.StartsWith("~$")) { continue }
    Write-Host "Processing $($file.Name)..."
    try {
        # LoadFile expects strict RTF or plain text.
        $rtf.LoadFile($file.FullName, [System.Windows.Forms.RichTextBoxStreamType]::RichText)
        $text = $rtf.Text
        $output += "`n`n================ FILE: $($file.Name) ================`n"
        $output += $text
        $output += "`n=====================================================`n"
        Write-Host "Success."
    } catch {
        Write-Host "Error: $_"
        $output += "`nError reading $($file.Name): $_`n"
    }
}

$outPath = "C:\geminiterminal2\extracted_text_ps.txt"
Write-Host "Saving to $outPath..."
[System.IO.File]::WriteAllText($outPath, $output)
Write-Host "Done."
