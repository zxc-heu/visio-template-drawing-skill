param(
  [Parameter(Mandatory=$true)][string]$InputVsdx,
  [Parameter(Mandatory=$true)][string]$OutputPng,
  [string]$PageName,
  [switch]$Visible
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $InputVsdx)) {
  throw "Input VSDX not found: $InputVsdx"
}

$outDir = Split-Path -Parent $OutputPng
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Path $outDir | Out-Null
}

$visio = New-Object -ComObject Visio.Application
$visio.Visible = [bool]$Visible

try {
  $doc = $visio.Documents.Open($InputVsdx)
  if ($PageName) {
    $page = $doc.Pages.ItemU($PageName)
  } else {
    $page = $visio.ActivePage
  }
  $visio.ActiveWindow.Page = $page
  $page.Export($OutputPng)
  $doc.Close()
  Write-Output "Exported preview: $OutputPng"
}
finally {
  if (-not $Visible) {
    $visio.Quit()
  }
}
