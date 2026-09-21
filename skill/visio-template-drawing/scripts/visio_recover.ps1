param([Parameter(Mandatory=$true)][string]$RunDirectory, [switch]$Recover)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\visio_session.ps1"
$result = Repair-VisioOrphan -RunDirectory ([IO.Path]::GetFullPath($RunDirectory)) -Recover:$Recover
$result | ConvertTo-Json -Depth 8
if ($Recover) { Write-VisioJson (Join-Path $RunDirectory 'recovery.manual.json') $result }
