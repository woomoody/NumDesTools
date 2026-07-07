param(
    [Parameter(Mandatory = $true)]
    [string] $RepoRoot
)

$ErrorActionPreference = 'Stop'

function Invoke-Step {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Label,
        [Parameter(Mandatory = $true)]
        [string[]] $Command
    )

    Write-Host "pre-commit: $Label"
    & dotnet @Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Label failed with exit code $LASTEXITCODE"
    }
}

Set-Location -LiteralPath $RepoRoot

Invoke-Step -Label 'build' -Command @('build', 'NumDesTools.sln', '-c', 'Debug')
Invoke-Step -Label 'test' -Command @('test', 'NumDesTools.Tests\NumDesTools.Tests.csproj', '-c', 'Debug', '--no-build')

Write-Host 'pre-commit: build and tests passed'
