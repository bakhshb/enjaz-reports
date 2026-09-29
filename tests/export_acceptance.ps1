param([Parameter(Mandatory=$true)][string]$ReportRoot)
$ErrorActionPreference='Stop'
# Match the Adobe PDFMaker route used by the approved reference PDFs.
$maker=New-Object -ComObject PDFMakerAPI.PDFMakerApp
$settings=New-Object -ComObject PDFMakerAPI.ConversionSettings
[void]$settings.LoadDefaultSettings()
$results=@()
$log=Join-Path $ReportRoot 'adobe-results.json'
if(Test-Path -LiteralPath $log){$results=@(Get-Content -LiteralPath $log -Raw | ConvertFrom-Json)}
foreach($deck in Get-ChildItem -LiteralPath $ReportRoot -Filter '*.pptx' -Recurse) {
    $before=(Get-FileHash -LiteralPath $deck.FullName -Algorithm SHA256).Hash
    $pdf=[IO.Path]::ChangeExtension($deck.FullName,'.pdf')
    $previous=@($results | Where-Object {$_.file -eq $deck.FullName -and $_.sha256 -eq $before -and $_.pdfSha256 -and (Test-Path -LiteralPath $pdf) -and $_.pdfSha256 -eq (Get-FileHash -LiteralPath $pdf -Algorithm SHA256).Hash})
    if($previous.Count){continue}
    # Publish only a complete export; PDFMaker can return a transient -8 in a batch.
    $temporary=Join-Path $deck.Directory.FullName ("export-"+[guid]::NewGuid().ToString()+'.pdf')
    for($attempt=1;$attempt -le 2;$attempt++) {
        $status=$maker.CreatePDF($deck.FullName,$temporary,$settings,$true,$false,$false,0)
        if($status -eq 0){break}
        if($attempt -eq 1) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($settings)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($maker)
            Start-Sleep -Seconds 2
            $maker=New-Object -ComObject PDFMakerAPI.PDFMakerApp
            $settings=New-Object -ComObject PDFMakerAPI.ConversionSettings
            [void]$settings.LoadDefaultSettings()
            $temporary=Join-Path $deck.Directory.FullName ("export-"+[guid]::NewGuid().ToString()+'.pdf')
        }
    }
    if($status -ne 0 -or -not (Test-Path -LiteralPath $temporary)){throw "PDFMaker failed ($status): $($deck.FullName)"}
    $after=(Get-FileHash -LiteralPath $deck.FullName -Algorithm SHA256).Hash
    if($before -ne $after){throw "Source deck changed during export: $($deck.FullName)"}
    Move-Item -LiteralPath $temporary -Destination $pdf -Force
    $results=@($results | Where-Object {$_.file -ne $deck.FullName})
    $results+=@{file=$deck.FullName;pdf=$pdf;sha256=$after;pdfSha256=(Get-FileHash -LiteralPath $pdf -Algorithm SHA256).Hash;exporter='Adobe PDFMaker';exported=$true}
    $results | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $log -Encoding utf8
    Write-Output "EXPORTED $($deck.Directory.Name) / $($deck.Name)"
}
