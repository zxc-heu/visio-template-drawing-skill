param([switch]$Visible, [string]$RunDirectory, [ValidateRange(5,86400)][int]$TimeoutSeconds = 60)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\visio_session.ps1"
Invoke-VisioJob -Request @{action='probe';visible=[bool]$Visible;keep_open=$false} -RunDirectory $RunDirectory -TimeoutSeconds $TimeoutSeconds
