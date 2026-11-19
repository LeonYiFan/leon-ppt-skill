param(
  [Parameter(Mandatory=$true)][string]$PptxPath,
  [Parameter(Mandatory=$true)][string]$OutputDir,
  [string]$FitManifest,
  [int[]]$SlideNumbers
)
$ErrorActionPreference='Stop'
$source=(Resolve-Path -LiteralPath $PptxPath).Path
$target=[IO.Path]::GetFullPath($OutputDir)
if(Test-Path -LiteralPath $target){
  if(@(Get-ChildItem -LiteralPath $target -Force).Count -gt 0){throw "Use a new empty render directory: $target"}
}
New-Item -ItemType Directory -Force -Path $target | Out-Null

function New-Rect([double]$Left,[double]$Top,[double]$Width,[double]$Height) {
  return [pscustomobject]@{Left=$Left;Top=$Top;Right=$Left+$Width;Bottom=$Top+$Height}
}

function Get-ShapeMeasurements($Shape,[bool]$AncestorRotated=$false,[string]$Cell='',[int]$TableId=0) {
  $angle = (([double]$Shape.Rotation % 360) + 360) % 360
  $rotated = $AncestorRotated -or ([Math]::Min($angle,360-$angle) -gt 0.01)
  if ($Shape.Visible -eq 0) {return}
  $box = New-Rect -Left $Shape.Left -Top $Shape.Top -Width $Shape.Width -Height $Shape.Height
  $text = $null
  $hasText = $Shape.HasTextFrame -eq -1 -and $Shape.TextFrame.HasText -eq -1
  if ($hasText -and !$rotated) {
    $frame = $Shape.TextFrame2
    $range = $frame.TextRange
    if (![string]::IsNullOrWhiteSpace($range.Text)) {
      if ($Cell) {
        # Cell range positions use PowerPoint's local anchor. Compare the measured spans with inner cell dimensions.
        $box = New-Rect 0 0 ($Shape.Width-$frame.MarginLeft-$frame.MarginRight) ($Shape.Height-$frame.MarginTop-$frame.MarginBottom)
        $text = New-Rect 0 0 $range.BoundWidth $range.BoundHeight
      } else {
        $text = New-Rect $range.BoundLeft $range.BoundTop $range.BoundWidth $range.BoundHeight
      }
    }
  }
  $id = if ($Cell) {$TableId} else {[int]$Shape.Id}
  [pscustomobject]@{Id=$id;Name=$Shape.Name;Box=$box;Text=$text;Rotated=$rotated;HasText=$hasText;Cell=$Cell}
  if ($Shape.Type -eq 6) {
    foreach ($child in $Shape.GroupItems) {Get-ShapeMeasurements $child $rotated}
  }
  if ($Shape.HasTable -eq -1) {
    $table = $Shape.Table
    for ($row=1; $row -le $table.Rows.Count; $row++) {
      for ($col=1; $col -le $table.Columns.Count; $col++) {
        Get-ShapeMeasurements $table.Cell($row,$col).Shape $rotated "[$row,$col]" $Shape.Id
      }
    }
  }
}

function Test-WithinRect($Entry,$Allowed,[int]$SlideNumber,[string]$Kind,$Container=$null) {
  # Small tolerance accommodates glyph overhang and differences in Office font metrics.
  $tolerance = 2.0
  $outside = [ordered]@{}
  foreach ($edge in 'Left','Right','Top','Bottom') {
    $sign = if ($edge -in 'Left','Top') {-1} else {1}
    $gap = ($Entry.Text.$edge - $Allowed.$edge) * $sign
    if ($gap -gt $tolerance) {$outside[$edge]=[Math]::Round($gap,2)}
  }
  if ($outside.Count -gt 0) {
    $axes = @(if ($outside.Contains('Left') -or $outside.Contains('Right')) {'horizontal'}
              if ($outside.Contains('Top') -or $outside.Contains('Bottom')) {'vertical'})
    [pscustomobject]@{
      Slide=$SlideNumber; Kind=$Kind; Shape=$Entry.Name; ShapeId=$Entry.Id; Cell=$Entry.Cell
      Container=$Container.Name; ContainerId=$Container.Id; Axes=$axes
      OverflowPt=$outside; BoundsPt=$Entry.Text; AllowedPt=$Allowed
    }
  }
}

function Get-ContainerChecks([string]$Path,[int]$PageCount) {
  $bySlide = @{}
  if (!$Path) {return $bySlide}
  $manifest = Get-Content -Raw -Encoding UTF8 -LiteralPath $Path | ConvertFrom-Json
  if ($manifest.slides -isnot [pscustomobject]) {throw 'Fit manifest requires a slides object'}
  foreach ($property in $manifest.slides.PSObject.Properties) {
    $page = 0
    if ($property.Name -notmatch '^[1-9][0-9]*$' -or ![int]::TryParse($property.Name,[ref]$page) -or $page -gt $PageCount) {
      throw "Invalid fit manifest slide: $($property.Name)"
    }
    $entries = $property.Value.containers
    if ($entries -isnot [array]) {throw "Fit manifest slide $page requires a containers array"}
    $children = @{}
    foreach ($entry in $entries) {
      $child = 0
      $parent = 0
      if (![int]::TryParse([string]$entry.shape_id,[ref]$child) -or $child -le 0 -or
          ![int]::TryParse([string]$entry.container_id,[ref]$parent) -or $parent -le 0 -or
          $child -eq $parent -or $children.ContainsKey($child)) {
        throw "Invalid or duplicate fit manifest shape mapping on slide $page"
      }
      $children[$child] = $true
      if ($entry.inset_pt -isnot [array] -or $entry.inset_pt.Count -ne 4) {
        throw "Fit manifest slide $page shape $child requires four insets [left, top, right, bottom] in points"
      }
      foreach ($inset in $entry.inset_pt) {
        if (($inset -isnot [ValueType]) -or $inset -is [bool] -or
            [double]::IsNaN([double]$inset) -or [double]::IsInfinity([double]$inset) -or $inset -lt 0) {
          throw "Invalid fit manifest inset on slide $page shape $child"
        }
      }
    }
    foreach ($slot in $property.Value.shapes) {
      if ($slot.container_required -eq $true -and !$children.ContainsKey([int]$slot.shape_id)) {
        throw "Fit manifest slide $page shape $($slot.shape_id) needs a container binding"
      }
    }
    $bySlide[$page] = $entries
  }
  return $bySlide
}

# A separate filename keeps COM from returning an already-open presentation.
$copy=Join-Path ([IO.Path]::GetTempPath()) ('leon-ppt-render-' + [guid]::NewGuid().ToString('N') + '.pptx')
$app=$null
$ownsDeck=$false
$deck=$null
$issues=[System.Collections.Generic.List[object]]::new()
$pagesChecked=[System.Collections.Generic.List[int]]::new()
$textChecked=0
$rotatedSkipped=0
$containersChecked=0
$failure=$null
$count=0
try {
  [IO.File]::Copy($source,$copy)
  $app=New-Object -ComObject PowerPoint.Application
  $deck=$app.Presentations.Open($copy,-1,0,0)
  $ownsDeck=$deck.FullName -eq $copy
  if(!$ownsDeck){throw 'PowerPoint returned a presentation other than the render copy'}
  $count=$deck.Slides.Count
  $deck.Export($target,'PNG',1620,912)
  $slides = New-Rect 0 0 $deck.PageSetup.SlideWidth $deck.PageSetup.SlideHeight
  $checks = Get-ContainerChecks $FitManifest $count
  foreach ($page in $SlideNumbers) {
    if ($page -lt 1 -or $page -gt $count) {throw "Slide number outside presentation: $page"}
  }
  foreach ($slide in $deck.Slides) {
    $page = [int]$slide.SlideIndex
    if ($SlideNumbers -and $page -notin $SlideNumbers) {continue}
    $byId = @{}
    foreach ($shape in $slide.Shapes) {
      foreach ($item in (Get-ShapeMeasurements $shape)) {
        if (!$item.Cell) {$byId[$item.Id]=$item}
        if ($item.Rotated -and $item.HasText) {$rotatedSkipped++}
        if (!$item.Text) {continue}
        $textChecked++
        $kind = if ($item.Cell) {'table-cell'} else {'text-frame'}
        $boxIssue = Test-WithinRect $item $item.Box $page $kind
        if ($boxIssue) {$issues.Add($boxIssue)}
        if (!$item.Cell) {
          $pageIssue = Test-WithinRect $item $slides $page 'slide'
          if ($pageIssue) {$issues.Add($pageIssue)}
        }
      }
    }
    foreach ($check in $checks[$page]) {
      $child = $byId[[int]$check.shape_id]
      $parent = $byId[[int]$check.container_id]
      if (!$child -or !$parent -or !$child.Text -or $parent.Rotated) {
        throw "Fit manifest slide $page cannot measure shape $($check.shape_id) inside container $($check.container_id)"
      }
      $inset = $check.inset_pt
      $safe = [pscustomobject]@{
        Left=$parent.Box.Left+$inset[0];Top=$parent.Box.Top+$inset[1]
        Right=$parent.Box.Right-$inset[2];Bottom=$parent.Box.Bottom-$inset[3]
      }
      if ($safe.Right -le $safe.Left -or $safe.Bottom -le $safe.Top) {
        throw "Fit manifest slide $page exhausts container $($parent.Id) with insets"
      }
      $containersChecked++
      $containerIssue = Test-WithinRect $child $safe $page 'container' $parent
      if ($containerIssue) {$issues.Add($containerIssue)}
    }
    $pagesChecked.Add($page)
  }
} catch {
  $failure=$_
} finally {
  try {
    if($ownsDeck){$deck.Close()}
  } finally {
    if($null -ne $deck){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck)}
    if($null -ne $app){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
    [IO.File]::Delete($copy)
  }
}

ConvertTo-Json -InputObject $issues.ToArray() -Depth 7 | Set-Content -LiteralPath (Join-Path $target 'text-fit.json') -Encoding utf8
[pscustomobject]@{
  Source=$source;SlideCount=$count;PngCount=@(Get-ChildItem -LiteralPath $target -Filter '*.PNG').Count
  SlidesChecked=$pagesChecked.ToArray();TextShapesChecked=$textChecked;RotatedTextShapesSkipped=$rotatedSkipped
  ContainerChecks=$containersChecked;TextFitIssues=$issues.Count;Error=if($failure){$failure.Exception.Message}else{$null}
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $target 'render-report.json') -Encoding utf8
if ($failure) {throw $failure}
if ($issues.Count) {throw "PowerPoint text fit failed with $($issues.Count) issue(s). See $target/text-fit.json."}
Write-Output "Rendered $count slides, checked $($pagesChecked.Count), and found no text fit issue. Inspect the PNGs before delivery."
