param([Parameter(Mandatory=$true)][string]$Report)
# Read PowerPoint's minimum heights from a disposable, content-populated probe.
$ErrorActionPreference='Stop'
$ppt=New-Object -ComObject PowerPoint.Application
$doc=$null
$result=@{}
try {
 $doc=$ppt.Presentations.Open((Resolve-Path -LiteralPath $Report).Path,-1,0,0)
 foreach($slide in $doc.Slides) {
  foreach($shape in $slide.Shapes) {
   if($shape.HasTable -ne -1){continue}
   $heights=@()
   foreach($row in $shape.Table.Rows){$heights+=@([double]$row.Height)}
   $result[$shape.Name]=$heights
  }
 }
} finally {
 if($doc){$doc.Close()}
 if($ppt.Presentations.Count -eq 0){$ppt.Quit()}
}
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath ([IO.Path]::ChangeExtension($Report,'.json')) -Encoding utf8
