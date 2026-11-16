# Opens each .docx in Word, prints "<pages><TAB><path>", and (with -Pdf) saves a PDF next to it.
# Used by run_issue.py to chain page numbers from one paper to the next.
param(
    [Parameter(Mandatory = $true)][string[]]$Paths,
    [switch]$Pdf
)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($p in $Paths) {
        $full = (Resolve-Path -LiteralPath $p).Path
        try {
            # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ..., OpenAndRepair = false)
            $doc = $word.Documents.Open($full, $false, $true, $false, "", "", $false, "", "", 0, 0, $false, $false)
            $doc.Repaginate()
            $pages = $doc.ComputeStatistics(2)
            if ($Pdf) {
                $out = [System.IO.Path]::ChangeExtension($full, ".pdf")
                # Item 0 = document content only, so Word's comment balloons are left out of the PDF
                $doc.ExportAsFixedFormat($out, 17, $false, 0, 0, 0, 0, 0)
            }
            $doc.Close($false)
            Write-Output ("{0}`t{1}" -f $pages, $full)
        } catch {
            Write-Output ("ERROR`t{0}`t{1}" -f $full, $_.Exception.Message)
        }
    }
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
