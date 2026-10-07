# Builds everything and runs the tests.
$ErrorActionPreference = 'Stop'
$sln = Join-Path (Split-Path $PSScriptRoot) 'ProjectPA.sln'
dotnet build $sln -c Release -v minimal --nologo
if ($LASTEXITCODE) { exit $LASTEXITCODE }
dotnet test $sln -c Release --no-build --nologo -v minimal
exit $LASTEXITCODE
