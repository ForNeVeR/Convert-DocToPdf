param (
    [string]$InputPath
)

$ErrorActionPreference = 'Stop'

$wdFormatPdf = 17
$InputPath = Resolve-Path $InputPath
$OutputPath = [IO.Path]::ChangeExtension($InputPath, 'pdf')

Write-Output "Saving $InputPath to $OutputPath..."
$word = New-Object -ComObject 'Word.Application'
try {
    $word.Visible = $false
    $document = $word.Documents.Open($InputPath)
    $document.SaveAs($OutputPath, $wdFormatPdf)
    $document.Close()
} finally {
    $word.Quit()
}
