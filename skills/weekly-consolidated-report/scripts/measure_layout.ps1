param(
 [Parameter(Mandatory=$true)][string]$Report,
 [Parameter(Mandatory=$true)][string]$Tasks,
 [Parameter(Mandatory=$true)][string]$Suhail,
 [Parameter(Mandatory=$true)][string]$Transactions,
 [string]$Template=(Join-Path $PSScriptRoot '../assets/weekly-report-master.pptx')
)
# Read native row heights after rendering reveals a layout correction is needed.
# No changes are made to the presentation, inputs, or template.
$ErrorActionPreference='Stop'
$hashes=@{}
foreach($entry in @{report=$Report;tasks=$Tasks;suhail=$Suhail;transactions=$Transactions;template=$Template}.GetEnumerator()) {
 $hashes[$entry.Key]=(Get-FileHash -LiteralPath $entry.Value -Algorithm SHA256).Hash
}
$ppt=New-Object -ComObject PowerPoint.Application
$doc=$null;$tables=@{}
try {
 $doc=$ppt.Presentations.Open((Resolve-Path -LiteralPath $Report).Path,-1,0,0)
 foreach($slide in $doc.Slides) {
  foreach($shape in @($slide.Shapes | Where-Object {$_.HasTable -eq -1})) {
   $table=$shape.Table
   $title=$table.Cell(1,1).Shape.TextFrame.TextRange.Text.Trim()
   if(-not $tables.ContainsKey($title)) {
    $tables[$title]=@{header=@([double]$table.Rows.Item(1).Height,[double]$table.Rows.Item(2).Height);body=@()}
   }
   for($r=3;$r -le $table.Rows.Count;$r++){$tables[$title].body+=@([double]$table.Rows.Item($r).Height)}
  }
 }
} finally {
 if($doc){$doc.Close()}
 if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
}
if($hashes.report -ne (Get-FileHash -LiteralPath $Report).Hash){throw 'Measuring changed the source presentation'}
$output=[IO.Path]::ChangeExtension($Report,'.layout.json')
@{hashes=$hashes;tables=$tables;visualAcceptance='pending'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $output -Encoding utf8
Write-Output $output
