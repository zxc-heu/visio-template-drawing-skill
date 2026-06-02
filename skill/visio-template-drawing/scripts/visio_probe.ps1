param(
  [switch]$Visible
)

$ErrorActionPreference = "Stop"

try {
  $visio = New-Object -ComObject Visio.Application
  $visio.Visible = [bool]$Visible
  $version = $visio.Version
  $caption = $visio.Caption
  Write-Output "Visio COM available"
  Write-Output "Version: $version"
  Write-Output "Caption: $caption"
  if (-not $Visible) {
    $visio.Quit()
  }
  exit 0
}
catch {
  Write-Error "Visio COM is not available: $($_.Exception.Message)"
  exit 1
}
