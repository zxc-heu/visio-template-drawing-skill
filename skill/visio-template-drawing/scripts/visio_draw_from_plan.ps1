param(
  [Parameter(Mandatory=$true)][string[]]$PlanJson,
  [string]$TemplateVsdx,
  [Parameter(Mandatory=$true)][string]$OutputVsdx,
  [switch]$Visible,
  [switch]$KeepOpen,
  [string]$PreviewDirectory,
  [string]$RunDirectory,
  [ValidateRange(5,86400)][int]$TimeoutSeconds = 300
)
$ErrorActionPreference = "Stop"
. "$PSScriptRoot\visio_session.ps1"
$plans=@($PlanJson | ForEach-Object { (Resolve-Path -LiteralPath $_).Path })
$output=[IO.Path]::GetFullPath($OutputVsdx)
if (Test-Path -LiteralPath $output) { throw "Output already exists; choose a new filename: $output" }
$template=if ($TemplateVsdx) { (Resolve-Path -LiteralPath $TemplateVsdx).Path } else { $null }
if ($template -and $template -ieq $output) { throw 'Output must differ from template.' }
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($output))
$preview=if ($PreviewDirectory) { [IO.Path]::GetFullPath($PreviewDirectory) } else { $null }
Invoke-VisioJob -Request @{action='draw';plans=$plans;template=$template;output=$output;visible=[bool]$Visible;keep_open=[bool]$KeepOpen;preview=$preview} -RunDirectory $RunDirectory -TimeoutSeconds $TimeoutSeconds
