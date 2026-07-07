param(
    [string] $RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'

Set-Location -LiteralPath $RepoRoot
git config --local core.hooksPath .githooks
Write-Host "Configured core.hooksPath=.githooks for $RepoRoot"
