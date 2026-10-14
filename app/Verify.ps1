param([Parameter(Mandatory=$true)][string]$OutputFile)
$ErrorActionPreference='Stop'
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'WordDefault.ps1') -LibraryOnly
@(Get-DefaultRows) | Export-Clixml -LiteralPath $OutputFile
