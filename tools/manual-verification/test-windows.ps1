param(
    [switch]$SkipGenerate
)

$ErrorActionPreference = "Stop"

$Phase11Dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Resolve-Path (Join-Path $Phase11Dir "..\..")
Set-Location $RepoRoot

if (-not $SkipGenerate) {
    Write-Host "==> Generating manual verification documents"
    dotnet run --project "tools/manual-verification/DocumentGenerator/DocumentGenerator.csproj"
}

Write-Host "==> Running Windows Microsoft Office evidence"
python "tools/manual-verification/scripts/run_windows_office_evidence.py"

Write-Host "==> Windows manual verification complete"
