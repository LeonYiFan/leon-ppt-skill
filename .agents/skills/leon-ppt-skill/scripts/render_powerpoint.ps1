param(
  [Parameter(Mandatory=$true)][string]$PptxPath,
  [Parameter(Mandatory=$true)][string]$OutputDir
)
$ErrorActionPreference='Stop'
$source=(Resolve-Path -LiteralPath $PptxPath).Path
$target=[IO.Path]::GetFullPath($OutputDir)
if(Test-Path -LiteralPath $target){
  if(@(Get-ChildItem -LiteralPath $target -Force).Count -gt 0){throw "Use a new empty render directory: $target"}
}
New-Item -ItemType Directory -Force -Path $target | Out-Null
# A separate filename keeps COM from returning an already-open presentation.
$copy=Join-Path ([IO.Path]::GetTempPath()) ('leon-ppt-render-' + [guid]::NewGuid().ToString('N') + '.pptx')
$app=$null
$ownsDeck=$false
$deck=$null
$issues=[System.Collections.Generic.List[object]]::new()
try {
  [IO.File]::Copy($source,$copy)
  $app=New-Object -ComObject PowerPoint.Application
  $deck=$app.Presentations.Open($copy,-1,0,0)
  $ownsDeck=$deck.FullName -eq $copy
  if(!$ownsDeck){throw 'PowerPoint returned a presentation other than the render copy'}
  $count=$deck.Slides.Count
  $deck.Export($target,'PNG',1620,912)
  foreach($slide in $deck.Slides){foreach($shape in $slide.Shapes){
    if($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1){
      if($shape.TextFrame2.TextRange.BoundHeight -gt ($shape.Height+2)){
        $issues.Add([pscustomobject]@{Slide=$slide.SlideIndex;Shape=$shape.Name;Text=$shape.TextFrame.TextRange.Text;Height=$shape.Height;BoundHeight=$shape.TextFrame2.TextRange.BoundHeight})
      }
    }
    if($shape.HasTable -eq -1){for($row=1;$row -le $shape.Table.Rows.Count;$row++){for($col=1;$col -le $shape.Table.Columns.Count;$col++){
      $cell=$shape.Table.Cell($row,$col).Shape
      if($cell.TextFrame2.TextRange.BoundHeight -gt ($cell.Height+2)){
        $issues.Add([pscustomobject]@{Slide=$slide.SlideIndex;Row=$row;Column=$col;Text=$cell.TextFrame.TextRange.Text;Height=$cell.Height;BoundHeight=$cell.TextFrame2.TextRange.BoundHeight})
      }
    }}}
  }}
  ConvertTo-Json -InputObject @($issues.ToArray()) -Depth 5 | Set-Content -LiteralPath (Join-Path $target 'text-fit.json') -Encoding utf8
  [pscustomobject]@{Source=$source;SlideCount=$count;PngCount=@(Get-ChildItem -LiteralPath $target -Filter '*.PNG').Count;TextFitIssues=$issues.Count} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $target 'render-report.json') -Encoding utf8
  Write-Output "Rendered $count slides. Text fit issues: $($issues.Count). Inspect every PNG before delivery."
} finally {
  try {
    if($ownsDeck){$deck.Close()}
  } finally {
    if($null -ne $deck){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck)}
    if($null -ne $app){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
    [IO.File]::Delete($copy)
  }
}
