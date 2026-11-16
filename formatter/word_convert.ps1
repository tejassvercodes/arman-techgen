# Converts a Word-readable file (.doc, .rtf, .odt ...) to .docx using Microsoft Word.
param(
    [Parameter(Mandatory = $true)][string]$InFile,
    [Parameter(Mandatory = $true)][string]$OutFile
)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open((Resolve-Path -LiteralPath $InFile).Path, $false, $true, $false, "", "", $false, "", "", 0, 0, $false, $false)
    $doc.SaveAs2([string]$OutFile, 16)
    $doc.Close($false)
    Write-Output "OK"
} catch {
    Write-Output ("ERROR`t{0}" -f $_.Exception.Message)
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
