param (
    [string]$docxPath,
    [string]$pdfPath
)

$wdFormatPDF = 17
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docxPath)
    $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
    $doc.Close([ref]0)
    Write-Host "SUCCESS: Exported PDF to $pdfPath"
} catch {
    Write-Error "ERROR: $_"
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
