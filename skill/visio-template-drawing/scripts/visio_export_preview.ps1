param(
  [Parameter(Mandatory=$true)][string]$InputVsdx,
  [Parameter(Mandatory=$true)][string]$OutputPng,
  [string]$PageName,
  [switch]$Visible,
  [switch]$KeepOpen,
  [string]$RunDirectory,
  [ValidateRange(5,86400)][int]$TimeoutSeconds = 300
)
$ErrorActionPreference = "Stop"
. "$PSScriptRoot\visio_session.ps1"
$inputPath=(Resolve-Path -LiteralPath $InputVsdx).Path
$output=[IO.Path]::GetFullPath($OutputPng)
if (Test-Path -LiteralPath $output) { throw "Export target already exists: $output" }
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($output))
Invoke-VisioJob -Request @{action='export';input=$inputPath;output=$output;preview=$output;page_name=$PageName;visible=[bool]$Visible;keep_open=[bool]$KeepOpen} -RunDirectory $RunDirectory -TimeoutSeconds $TimeoutSeconds
