param(
  [Parameter(Mandatory=$true)][string]$PlanJson,
  [Parameter(Mandatory=$true)][string]$TemplateVsdx,
  [Parameter(Mandatory=$true)][string]$OutputVsdx,
  [switch]$Visible
)

$ErrorActionPreference = "Stop"

function Convert-HexToVisioRgbFormula {
  param([string]$Hex)
  if ([string]::IsNullOrWhiteSpace($Hex)) { return $null }
  $value = $Hex.Trim().TrimStart("#")
  if ($value.Length -ne 6) { return $null }
  $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
  return "RGB($r,$g,$b)"
}

function Set-CellFormulaIfPresent {
  param($Shape, [string]$CellName, [string]$Formula)
  if ([string]::IsNullOrWhiteSpace($Formula)) { return }
  try {
    $Shape.CellsU($CellName).FormulaU = $Formula
  } catch {
    Write-Warning "Could not set $CellName on shape '$($Shape.NameU)': $($_.Exception.Message)"
  }
}

function Set-ShapeFillRecursive {
  param($Shape, [string]$FillFormula, [string]$LineFormula)
  if ($null -eq $Shape) { return }
  Set-CellFormulaIfPresent -Shape $Shape -CellName "FillForegnd" -Formula $FillFormula
  Set-CellFormulaIfPresent -Shape $Shape -CellName "LineColor" -Formula $LineFormula
  try {
    foreach ($child in $Shape.Shapes) {
      Set-ShapeFillRecursive -Shape $child -FillFormula $FillFormula -LineFormula $LineFormula
    }
  } catch {
  }
}

function Find-MasterByName {
  param($Application, [string]$MasterName)
  if ([string]::IsNullOrWhiteSpace($MasterName)) { return $null }
  foreach ($doc in $Application.Documents) {
    foreach ($master in $doc.Masters) {
      if ($master.NameU -eq $MasterName -or $master.Name -eq $MasterName) {
        return $master
      }
    }
  }
  return $null
}

function Find-FirstMasterByName {
  param($Application, [string[]]$MasterNames)
  foreach ($name in $MasterNames) {
    $master = Find-MasterByName -Application $Application -MasterName $name
    if ($master) { return $master }
  }
  return $null
}

function Get-VisioContentRoots {
  $candidates = @(
    (Join-Path $env:ProgramFiles "Microsoft Office\root\Office16\Visio Content"),
    (Join-Path ${env:ProgramFiles(x86)} "Microsoft Office\root\Office16\Visio Content"),
    (Join-Path $env:ProgramFiles "Microsoft Office\Office16\Visio Content"),
    (Join-Path ${env:ProgramFiles(x86)} "Microsoft Office\Office16\Visio Content")
  )
  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) {
      Get-Item -LiteralPath $candidate
    }
  }
}

function Find-VisioStencilFile {
  param([string]$NameOrRegex)
  if ([string]::IsNullOrWhiteSpace($NameOrRegex)) { return $null }
  if (Test-Path -LiteralPath $NameOrRegex) { return (Get-Item -LiteralPath $NameOrRegex) }
  foreach ($root in Get-VisioContentRoots) {
    $matches = Get-ChildItem -Path $root.FullName -Recurse -File -Filter "*.vss*" -ErrorAction SilentlyContinue |
      Where-Object { $_.Name -ieq $NameOrRegex -or $_.Name -match $NameOrRegex } |
      Sort-Object Name
    $first = @($matches | Select-Object -First 1)
    if ($first.Count -gt 0) { return $first[0] }
  }
  return $null
}

function Open-VisioStencilReadOnly {
  param($Application, [string]$StencilNameOrPath)
  $file = Find-VisioStencilFile -NameOrRegex $StencilNameOrPath
  if (-not $file) { return $null }
  foreach ($doc in $Application.Documents) {
    try {
      if ($doc.FullName -and ([System.IO.Path]::GetFullPath($doc.FullName) -ieq $file.FullName)) { return $doc }
    } catch {
    }
  }
  # 66 = hidden + read-only. This avoids visible stencil windows and save prompts.
  return $Application.Documents.OpenEx($file.FullName, 66)
}

function Open-CommonShapeStencils {
  param($Application)
  foreach ($fileName in @("BLOCK_M.VSSX", "BASIC_M.VSSX", "BASFLO_M.VSSX")) {
    try {
      Open-VisioStencilReadOnly -Application $Application -StencilNameOrPath $fileName | Out-Null
    } catch {
      Write-Warning "Could not open Visio stencil '$fileName': $($_.Exception.Message)"
    }
  }
}

function Find-VisioMasterDiscovery {
  param(
    $Application,
    [string]$Query,
    [string]$PreferredStencilRegex = "",
    [int]$MaxStencils = 16
  )
  if ([string]::IsNullOrWhiteSpace($Query)) { return $null }
  $files = @()
  foreach ($root in Get-VisioContentRoots) {
    $rootFiles = Get-ChildItem -Path $root.FullName -Recurse -File -Filter "*.vss*" -ErrorAction SilentlyContinue
    if ($PreferredStencilRegex) {
      $files += @($rootFiles | Where-Object { $_.Name -match $PreferredStencilRegex })
    } else {
      $files += @($rootFiles)
    }
  }
  foreach ($file in @($files | Sort-Object Name | Select-Object -First $MaxStencils)) {
    try {
      $stencil = Open-VisioStencilReadOnly -Application $Application -StencilNameOrPath $file.FullName
      foreach ($master in $stencil.Masters) {
        if ($master.NameU -match $Query -or $master.Name -match $Query) {
          return $master
        }
      }
    } catch {
    }
  }
  return $null
}

function Test-ContainsCjk {
  param([string]$Text)
  if ([string]::IsNullOrEmpty($Text)) { return $false }
  return [regex]::IsMatch($Text, "[\u4e00-\u9fff]")
}

function Get-DefaultFontFamily {
  param([string]$Text)
  if (Test-ContainsCjk -Text $Text) { return "SimSun" }
  return "Times New Roman"
}

function Get-VisioFontId {
  param($Document, [string]$FontFamily)
  if ([string]::IsNullOrWhiteSpace($FontFamily)) { return $null }
  try {
    return [int]$Document.Fonts.Item($FontFamily).ID
  } catch {
    return $null
  }
}

function Set-ShapeFontFamily {
  param($Document, $Shape, [string]$FontFamily)
  if ([string]::IsNullOrWhiteSpace($FontFamily)) { return }
  try {
    $Shape.CellsU("Char.Font").FormulaU = "FONT(`"$FontFamily`")"
  } catch {
    Write-Warning "Could not set font '$FontFamily' on shape '$($Shape.NameU)': $($_.Exception.Message)"
  }
}

function Apply-RichText {
  param(
    $Document,
    $Shape,
    $Spans,
    [string]$DefaultFontFamily,
    [double]$DefaultFontSize = 10
  )

  $spanList = @($Spans)
  if ($spanList.Count -eq 0) { return }

  $fullText = ""
  foreach ($span in $spanList) {
    if ($null -ne $span.text) {
      $fullText += [string]$span.text
    }
  }
  $Shape.Text = $fullText

  $offset = 0
  foreach ($span in $spanList) {
    $text = if ($null -ne $span.text) { [string]$span.text } else { "" }
    $length = $text.Length
    if ($length -le 0) { continue }

    $chars = $Shape.Characters
    $chars.Begin = $offset
    $chars.End = $offset + $length

    $fontFamily = if ($span.fontFamily) { [string]$span.fontFamily } else { $DefaultFontFamily }
    $fontId = Get-VisioFontId -Document $Document -FontFamily $fontFamily
    if ($null -ne $fontId) {
      try { $chars.CharProps(0) = $fontId } catch { Write-Warning "Could not set rich text font '$fontFamily' on shape '$($Shape.NameU)': $($_.Exception.Message)" }
    }

    $fontSize = if ($span.fontSize) { [double]$span.fontSize } elseif ($span.subscript -eq $true -or $span.superscript -eq $true) { [double]$DefaultFontSize * 0.65 } else { $DefaultFontSize }
    if ($fontSize -gt 0) {
      try { $chars.CharProps(7) = $fontSize } catch { Write-Warning "Could not set rich text size on shape '$($Shape.NameU)': $($_.Exception.Message)" }
    }

    if ($span.bold -eq $true -or $span.italic -eq $true) {
      $style = 0
      if ($span.bold -eq $true) { $style += 1 }
      if ($span.italic -eq $true) { $style += 2 }
      try { $chars.CharProps(2) = $style } catch { Write-Warning "Could not set rich text style on shape '$($Shape.NameU)': $($_.Exception.Message)" }
    }
    if ($span.subscript -eq $true) {
      try { $chars.CharProps(4) = 2 } catch { Write-Warning "Could not set rich text subscript on shape '$($Shape.NameU)': $($_.Exception.Message)" }
    } elseif ($span.superscript -eq $true) {
      try { $chars.CharProps(4) = 1 } catch { Write-Warning "Could not set rich text superscript on shape '$($Shape.NameU)': $($_.Exception.Message)" }
    }

    $offset += $length
  }
}

function Draw-PolygonByPoints {
  param($Page, [string]$Id, $Points, [string]$Fill, [string]$Line)
  $nums = @($Points) | ForEach-Object { [double]$_ }
  if ($nums.Count -lt 6) { return $null }
  if ($nums[0] -ne $nums[$nums.Count - 2] -or $nums[1] -ne $nums[$nums.Count - 1]) {
    $nums += $nums[0]
    $nums += $nums[1]
  }
  $shape = $Page.DrawPolyline(([double[]]$nums), 0)
  $shape.NameU = $Id
  Set-CellFormulaIfPresent -Shape $shape -CellName "FillPattern" -Formula "1"
  Set-CellFormulaIfPresent -Shape $shape -CellName "FillForegnd" -Formula (Convert-HexToVisioRgbFormula $Fill)
  if ([string]::IsNullOrWhiteSpace($Line)) {
    Set-CellFormulaIfPresent -Shape $shape -CellName "LinePattern" -Formula "0"
  } else {
    Set-CellFormulaIfPresent -Shape $shape -CellName "LineColor" -Formula (Convert-HexToVisioRgbFormula $Line)
  }
  return $shape
}

function Convert-CenterlineToRibbonPoints {
  param($Points, [double]$Width)
  $nums = @($Points) | ForEach-Object { [double]$_ }
  if ($nums.Count -lt 4 -or $Width -le 0) { return $nums }

  $left = New-Object System.Collections.Generic.List[double]
  $right = New-Object System.Collections.Generic.List[double]
  $half = $Width / 2.0
  $pointCount = [int]($nums.Count / 2)

  for ($i = 0; $i -lt $pointCount; $i++) {
    $x = $nums[$i * 2]
    $y = $nums[$i * 2 + 1]
    if ($i -eq 0) {
      $nx = $nums[2] - $x
      $ny = $nums[3] - $y
    } elseif ($i -eq ($pointCount - 1)) {
      $nx = $x - $nums[($i - 1) * 2]
      $ny = $y - $nums[($i - 1) * 2 + 1]
    } else {
      $nx = $nums[($i + 1) * 2] - $nums[($i - 1) * 2]
      $ny = $nums[($i + 1) * 2 + 1] - $nums[($i - 1) * 2 + 1]
    }
    $len = [Math]::Sqrt(($nx * $nx) + ($ny * $ny))
    if ($len -eq 0) {
      $px = 0
      $py = $half
    } else {
      $px = -$ny / $len * $half
      $py = $nx / $len * $half
    }
    $left.Add($x + $px)
    $left.Add($y + $py)
    $right.Insert(0, $y - $py)
    $right.Insert(0, $x - $px)
  }

  $result = New-Object System.Collections.Generic.List[double]
  foreach ($n in $left) { $result.Add($n) }
  foreach ($n in $right) { $result.Add($n) }
  return [double[]]$result
}

function Draw-RibbonByPoints {
  param($Page, [string]$Id, $Points, [string]$Fill, [string]$Line, [double]$Width = 0)
  $polygonPoints = if ($Width -gt 0) { Convert-CenterlineToRibbonPoints -Points $Points -Width $Width } else { $Points }
  $shape = Draw-PolygonByPoints -Page $Page -Id $Id -Points $polygonPoints -Fill $Fill -Line $Line
  if ($shape) {
    Set-CellFormulaIfPresent -Shape $shape -CellName "LineWeight" -Formula "0.75 pt"
  }
  return $shape
}

function Draw-Cuboid {
  param($Page, [string]$Id, [double]$X, [double]$Y, [double]$W, [double]$H, [string]$Fill, [string]$Line)
  $depthX = [Math]::Min(0.12, $W * 0.08)
  $depthY = [Math]::Min(0.08, $H * 0.20)
  $left = $X - ($W / 2)
  $right = $X + ($W / 2)
  $bottom = $Y - ($H / 2)
  $top = $Y + ($H / 2)
  $fillFormula = Convert-HexToVisioRgbFormula $Fill
  $lineFormula = Convert-HexToVisioRgbFormula $Line

  $front = $Page.DrawRectangle($left, $bottom, $right, $top)
  $front.NameU = $Id
  Set-CellFormulaIfPresent -Shape $front -CellName "FillForegnd" -Formula $fillFormula
  Set-CellFormulaIfPresent -Shape $front -CellName "LineColor" -Formula $lineFormula

  $side = $Page.DrawPolyline(([double[]]@($right,$bottom, $right+$depthX,$bottom+$depthY, $right+$depthX,$top+$depthY, $right,$top, $right,$bottom)), 0)
  $side.NameU = $Id + "_side"
  Set-CellFormulaIfPresent -Shape $side -CellName "FillPattern" -Formula "1"
  Set-CellFormulaIfPresent -Shape $side -CellName "FillForegnd" -Formula $fillFormula
  Set-CellFormulaIfPresent -Shape $side -CellName "LineColor" -Formula $lineFormula

  return $front
}

function Draw-PolylineByPoints {
  param($Page, [string]$Id, $Points, [string]$Line, [bool]$Arrow, [bool]$Dashed, [bool]$Rounded)
  $nums = @($Points) | ForEach-Object { [double]$_ }
  if ($nums.Count -lt 4) { return $null }
  $shape = $Page.DrawPolyline(([double[]]$nums), 0)
  $shape.NameU = $Id
  Set-CellFormulaIfPresent -Shape $shape -CellName "LineColor" -Formula (Convert-HexToVisioRgbFormula $Line)
  if ($Arrow) { Set-CellFormulaIfPresent -Shape $shape -CellName "EndArrow" -Formula "4" }
  if ($Dashed) { Set-CellFormulaIfPresent -Shape $shape -CellName "LinePattern" -Formula "2" }
  if ($Rounded) { Set-CellFormulaIfPresent -Shape $shape -CellName "Rounding" -Formula "0.18 in" }
  return $shape
}

function Draw-SplineByPoints {
  param($Page, [string]$Id, $Points, [string]$Line, [bool]$Arrow, [bool]$Dashed)
  $nums = @($Points) | ForEach-Object { [double]$_ }
  if ($nums.Count -lt 4) { return $null }
  try {
    $shape = $Page.DrawSpline(([double[]]$nums), 0.05, 0)
  } catch {
    Write-Warning "DrawSpline failed for '$Id'; falling back to editable rounded polyline: $($_.Exception.Message)"
    $shape = $Page.DrawPolyline(([double[]]$nums), 0)
    Set-CellFormulaIfPresent -Shape $shape -CellName "Rounding" -Formula "0.18 in"
  }
  $shape.NameU = $Id
  Set-CellFormulaIfPresent -Shape $shape -CellName "LineColor" -Formula (Convert-HexToVisioRgbFormula $Line)
  if ($Arrow) { Set-CellFormulaIfPresent -Shape $shape -CellName "EndArrow" -Formula "4" }
  if ($Dashed) { Set-CellFormulaIfPresent -Shape $shape -CellName "LinePattern" -Formula "2" }
  return $shape
}

function Draw-BezierByPoints {
  param($Page, [string]$Id, $Points, [string]$Line, [bool]$Arrow, [bool]$Dashed, [int]$Degree = 3)
  $nums = @($Points) | ForEach-Object { [double]$_ }
  if ($nums.Count -lt 6) { return $null }
  try {
    $shape = $Page.DrawBezier(([double[]]$nums), $Degree, 8)
  } catch {
    Write-Warning "DrawBezier failed for '$Id'; falling back to DrawSpline: $($_.Exception.Message)"
    $shape = $Page.DrawSpline(([double[]]$nums), 0.04, 8)
  }
  $shape.NameU = $Id
  Set-CellFormulaIfPresent -Shape $shape -CellName "LineColor" -Formula (Convert-HexToVisioRgbFormula $Line)
  if ($Arrow) { Set-CellFormulaIfPresent -Shape $shape -CellName "EndArrow" -Formula "4" }
  if ($Dashed) { Set-CellFormulaIfPresent -Shape $shape -CellName "LinePattern" -Formula "2" }
  return $shape
}

function Get-DynamicConnectorMaster {
  param($Application)
  Open-CommonShapeStencils -Application $Application
  $master = Find-FirstMasterByName -Application $Application -MasterNames @("Dynamic connector", "Dynamic Connector")
  if (-not $master) {
    $master = Find-VisioMasterDiscovery -Application $Application -Query "Dynamic connector|Dynamic Connector" -PreferredStencilRegex "BASFLO|BASIC|BLOCK"
  }
  return $master
}

if (-not (Test-Path -LiteralPath $PlanJson)) {
  throw "Plan JSON not found: $PlanJson"
}
if (-not (Test-Path -LiteralPath $TemplateVsdx)) {
  throw "Template VSDX not found: $TemplateVsdx"
}
if ((Resolve-Path -LiteralPath $TemplateVsdx).Path -eq $OutputVsdx) {
  throw "OutputVsdx must not be the same path as TemplateVsdx"
}

$plan = Get-Content -LiteralPath $PlanJson -Raw -Encoding UTF8 | ConvertFrom-Json
$outDir = Split-Path -Parent $OutputVsdx
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Path $outDir | Out-Null
}

$visio = New-Object -ComObject Visio.Application
$visio.Visible = [bool]$Visible
$shapeById = @{}
$commonStencilsOpened = $false

try {
  $doc = $visio.Documents.Open($TemplateVsdx)
  $pageName = if ($plan.page.name) { [string]$plan.page.name } else { "Generated Diagram" }

  $page = $null
  foreach ($candidate in $doc.Pages) {
    if ($candidate.NameU -eq $pageName -or $candidate.Name -eq $pageName) {
      $page = $candidate
      break
    }
  }
  if (-not $page) {
    $page = $doc.Pages.Add()
    $page.Name = $pageName
  }
  $visio.ActiveWindow.Page = $page

  if ($plan.page.width) {
    $page.PageSheet.CellsU("PageWidth").ResultIU = [double]$plan.page.width
  }
  if ($plan.page.height) {
    $page.PageSheet.CellsU("PageHeight").ResultIU = [double]$plan.page.height
  }
  $defaultConnectorZOrder = if ($plan.defaults -and $plan.defaults.connectorZOrder) { ([string]$plan.defaults.connectorZOrder).ToLowerInvariant() } else { "" }

  foreach ($item in $plan.shapes) {
    $id = [string]$item.id
    if ([string]::IsNullOrWhiteSpace($id)) {
      throw "Every shape must have an id"
    }

    $x = [double]$item.x
    $y = [double]$item.y
    $w = if ($item.width) { [double]$item.width } else { 1.5 }
    $h = if ($item.height) { [double]$item.height } else { 0.6 }
    $shape = $null
    if ($item.stencil) {
      try {
        Open-VisioStencilReadOnly -Application $visio -StencilNameOrPath ([string]$item.stencil) | Out-Null
      } catch {
        Write-Warning "Could not open requested stencil '$($item.stencil)' for shape '$id': $($_.Exception.Message)"
      }
    }
    $master = Find-MasterByName -Application $visio -MasterName ([string]$item.master)
    if (-not $master -and $item.masterQuery) {
      $master = Find-VisioMasterDiscovery -Application $visio -Query ([string]$item.masterQuery) -PreferredStencilRegex ([string]$item.preferredStencilRegex)
      if ($master) {
        Write-Output "Discovered master for '$id': $($master.NameU)"
      }
    }

    if ($master) {
      $shape = $page.Drop($master, $x, $y)
      $shape.CellsU("Width").ResultIU = $w
      $shape.CellsU("Height").ResultIU = $h
    } else {
      $kind = if ($item.kind) { ([string]$item.kind).ToLowerInvariant() } else { "rectangle" }
      switch ($kind) {
        "ellipse" {
          $shape = $page.DrawOval($x - ($w / 2), $y - ($h / 2), $x + ($w / 2), $y + ($h / 2))
        }
        "diamond" {
          $diamondPoints = @(
            $x, ($y + ($h / 2)),
            ($x + ($w / 2)), $y,
            $x, ($y - ($h / 2)),
            ($x - ($w / 2)), $y
          )
          $shape = Draw-PolygonByPoints -Page $page -Id $id -Points $diamondPoints -Fill ([string]$item.fill) -Line ([string]$item.line)
        }
        "text" {
          $shape = $page.DrawRectangle($x - ($w / 2), $y - ($h / 2), $x + ($w / 2), $y + ($h / 2))
          Set-CellFormulaIfPresent -Shape $shape -CellName "LinePattern" -Formula "0"
          Set-CellFormulaIfPresent -Shape $shape -CellName "FillPattern" -Formula "0"
        }
        "image" {
          $imagePath = [string]$item.imagePath
          if ([string]::IsNullOrWhiteSpace($imagePath) -or -not (Test-Path -LiteralPath $imagePath)) {
            Write-Warning "Image path missing for shape '$id'; drawing placeholder instead."
            $shape = $page.DrawRectangle($x - ($w / 2), $y - ($h / 2), $x + ($w / 2), $y + ($h / 2))
            Set-CellFormulaIfPresent -Shape $shape -CellName "FillForegnd" -Formula "RGB(240,240,240)"
          } else {
            $shape = $page.Import($imagePath)
            $shape.CellsU("PinX").ResultIU = $x
            $shape.CellsU("PinY").ResultIU = $y
            $shape.CellsU("Width").ResultIU = $w
            $shape.CellsU("Height").ResultIU = $h
          }
        }
        "polygon" {
          $shape = Draw-PolygonByPoints -Page $page -Id $id -Points $item.points -Fill ([string]$item.fill) -Line ([string]$item.line)
          if (-not $shape) {
            Write-Warning "Polygon points missing for shape '$id'; drawing placeholder rectangle instead."
            $shape = $page.DrawRectangle($x - ($w / 2), $y - ($h / 2), $x + ($w / 2), $y + ($h / 2))
          }
        }
        "cuboid" {
          if (-not $commonStencilsOpened) {
            Open-CommonShapeStencils -Application $visio
            $commonStencilsOpened = $true
          }
          $cuboidMaster = Find-FirstMasterByName -Application $visio -MasterNames @("3-D box", "Cube", "Horizontal bar", "Square block")
          if ($cuboidMaster) {
            $shape = $page.Drop($cuboidMaster, $x, $y)
            $shape.CellsU("Width").ResultIU = $w
            $shape.CellsU("Height").ResultIU = $h
            Set-ShapeFillRecursive -Shape $shape -FillFormula (Convert-HexToVisioRgbFormula ([string]$item.fill)) -LineFormula (Convert-HexToVisioRgbFormula ([string]$item.line))
          } else {
            Write-Warning "No cuboid master found in common stencils; drawing editable fallback faces for '$id'."
            $shape = Draw-Cuboid -Page $page -Id $id -X $x -Y $y -W $w -H $h -Fill ([string]$item.fill) -Line ([string]$item.line)
          }
        }
        default {
          $shape = $page.DrawRectangle($x - ($w / 2), $y - ($h / 2), $x + ($w / 2), $y + ($h / 2))
          if ($kind -eq "rounded-rectangle" -or $item.rounded) {
            Set-CellFormulaIfPresent -Shape $shape -CellName "Rounding" -Formula "0.12 in"
          }
        }
      }
    }

    $shape.NameU = $id
    $hasRichText = ($null -ne $item.richText -and @($item.richText).Count -gt 0)
    if (-not $hasRichText -and $null -ne $item.text) { $shape.Text = [string]$item.text }
    $fontFamily = if ($item.fontFamily) { [string]$item.fontFamily } else { Get-DefaultFontFamily -Text ([string]$item.text) }
    Set-ShapeFontFamily -Document $doc -Shape $shape -FontFamily $fontFamily
    Set-CellFormulaIfPresent -Shape $shape -CellName "FillForegnd" -Formula (Convert-HexToVisioRgbFormula ([string]$item.fill))
    Set-CellFormulaIfPresent -Shape $shape -CellName "LineColor" -Formula (Convert-HexToVisioRgbFormula ([string]$item.line))
    Set-CellFormulaIfPresent -Shape $shape -CellName "Char.Color" -Formula (Convert-HexToVisioRgbFormula ([string]$item.fontColor))
    if ($item.lineWeight) {
      Set-CellFormulaIfPresent -Shape $shape -CellName "LineWeight" -Formula "$([double]$item.lineWeight) pt"
    }
    if ($item.fontSize) {
      Set-CellFormulaIfPresent -Shape $shape -CellName "Char.Size" -Formula "$([double]$item.fontSize) pt"
    }
    if ($item.bold -eq $true -or $item.italic -eq $true) {
      $style = 0
      if ($item.bold -eq $true) { $style += 1 }
      if ($item.italic -eq $true) { $style += 2 }
      Set-CellFormulaIfPresent -Shape $shape -CellName "Char.Style" -Formula ([string]$style)
    }
    if ($hasRichText) {
      $defaultSize = if ($item.fontSize) { [double]$item.fontSize } else { 10 }
      Apply-RichText -Document $doc -Shape $shape -Spans $item.richText -DefaultFontFamily $fontFamily -DefaultFontSize $defaultSize
    }
    $shapeById[$id] = $shape
  }

  foreach ($conn in $plan.connectors) {
    $fromId = [string]$conn.from
    $toId = [string]$conn.to
    if (-not $shapeById.ContainsKey($fromId) -or -not $shapeById.ContainsKey($toId)) {
      Write-Warning "Skipping connector '$($conn.id)' because endpoint is missing: $fromId -> $toId"
      continue
    }

    $fromShape = $shapeById[$fromId]
    $toShape = $shapeById[$toId]
    $connId = if ($conn.id) { [string]$conn.id } else { "connector_$fromId`_$toId" }
    $connKind = if ($conn.kind) { ([string]$conn.kind).ToLowerInvariant() } else { "dynamic" }
    $connector = $null
    if (($connKind -eq "ribbon" -or $connKind -eq "polygon") -and $conn.points) {
      $ribbonFill = if ($conn.fill) { [string]$conn.fill } else { [string]$conn.line }
      $ribbonWidth = if ($conn.width) { [double]$conn.width } else { 0 }
      $connector = Draw-RibbonByPoints -Page $page -Id $connId -Points $conn.points -Fill $ribbonFill -Line ([string]$conn.line) -Width $ribbonWidth
    } elseif ($connKind -eq "bezier" -and $conn.points) {
      $degree = if ($conn.degree) { [int]$conn.degree } else { 3 }
      $connector = Draw-BezierByPoints -Page $page -Id $connId -Points $conn.points -Line ([string]$conn.line) -Arrow ($conn.arrow -eq $true) -Dashed ($conn.dashed -eq $true) -Degree $degree
    } elseif ($connKind -eq "curve" -and $conn.points) {
      $connector = Draw-SplineByPoints -Page $page -Id $connId -Points $conn.points -Line ([string]$conn.line) -Arrow ($conn.arrow -eq $true) -Dashed ($conn.dashed -eq $true)
    } elseif ($connKind -eq "polyline" -and $conn.points) {
      $connector = Draw-PolylineByPoints -Page $page -Id $connId -Points $conn.points -Line ([string]$conn.line) -Arrow ($conn.arrow -eq $true) -Dashed ($conn.dashed -eq $true) -Rounded $false
    } elseif ($connKind -eq "straight") {
      $connector = $page.DrawLine($fromShape.CellsU("PinX").ResultIU, $fromShape.CellsU("PinY").ResultIU, $toShape.CellsU("PinX").ResultIU, $toShape.CellsU("PinY").ResultIU)
      $connector.NameU = $connId
    } else {
      $connectorMaster = Get-DynamicConnectorMaster -Application $visio
      if ($connectorMaster) {
        $connector = $page.Drop($connectorMaster, 0, 0)
      } else {
        $connector = $page.Drop($visio.ConnectorToolDataObject, 0, 0)
      }
      $connector.NameU = $connId
      $fromX = if ($conn.fromX -ne $null) { [double]$conn.fromX } else { 1.0 }
      $fromY = if ($conn.fromY -ne $null) { [double]$conn.fromY } else { 0.5 }
      $toX = if ($conn.toX -ne $null) { [double]$conn.toX } else { 0.0 }
      $toY = if ($conn.toY -ne $null) { [double]$conn.toY } else { 0.5 }
      try {
        $connector.CellsU("BeginX").GlueToPos($fromShape, $fromX, $fromY)
        $connector.CellsU("EndX").GlueToPos($toShape, $toX, $toY)
      } catch {
        Write-Warning "GlueToPos failed for connector '$connId'; falling back to center glue: $($_.Exception.Message)"
        $connector.CellsU("BeginX").GlueTo($fromShape.CellsU("PinX"))
        $connector.CellsU("EndX").GlueTo($toShape.CellsU("PinX"))
      }
      if ($connector.CellExistsU("ShapeRouteStyle", 0) -ne 0) {
        $connector.CellsU("ShapeRouteStyle").FormulaU = "0"
      }
      if ($connector.CellExistsU("ConLineRouteExt", 0) -ne 0) {
        $connector.CellsU("ConLineRouteExt").FormulaU = "0"
      }
    }
    if (-not $connector) { continue }
    if ($null -ne $conn.text) { $connector.Text = [string]$conn.text }
    $connFontFamily = if ($conn.fontFamily) { [string]$conn.fontFamily } else { Get-DefaultFontFamily -Text ([string]$conn.text) }
    Set-ShapeFontFamily -Document $doc -Shape $connector -FontFamily $connFontFamily
    Set-CellFormulaIfPresent -Shape $connector -CellName "LineColor" -Formula (Convert-HexToVisioRgbFormula ([string]$conn.line))
    if ($conn.lineWeight) {
      Set-CellFormulaIfPresent -Shape $connector -CellName "LineWeight" -Formula "$([double]$conn.lineWeight) pt"
    }
    if ($conn.arrow -eq $true) {
      Set-CellFormulaIfPresent -Shape $connector -CellName "EndArrow" -Formula "4"
    }
    if ($conn.bold -eq $true -or $conn.italic -eq $true) {
      $connStyle = 0
      if ($conn.bold -eq $true) { $connStyle += 1 }
      if ($conn.italic -eq $true) { $connStyle += 2 }
      Set-CellFormulaIfPresent -Shape $connector -CellName "Char.Style" -Formula ([string]$connStyle)
    }
    $zOrder = if ($conn.zOrder) { ([string]$conn.zOrder).ToLowerInvariant() } else { $defaultConnectorZOrder }
    if ($zOrder -eq "back") {
      try { $connector.SendToBack() } catch { Write-Warning "Could not send connector '$connId' to back: $($_.Exception.Message)" }
    } elseif ($zOrder -eq "front") {
      try { $connector.BringToFront() } catch { Write-Warning "Could not bring connector '$connId' to front: $($_.Exception.Message)" }
    }
  }

  $doc.SaveAs($OutputVsdx)
  Write-Output "Saved editable Visio drawing: $OutputVsdx"
}
finally {
  if (-not $Visible) {
    $visio.Quit()
  }
}
