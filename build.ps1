param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'dist\LimoPcBuildSheetManager.exe')
)

$ErrorActionPreference = 'Stop'
$compilerCandidates = @(
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
)
$compiler = $compilerCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $compiler) {
    throw 'The .NET Framework 4.x C# compiler was not found.'
}

$outputDirectory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

& $compiler /nologo /target:winexe /optimize+ /platform:anycpu `
    /reference:System.dll `
    /reference:System.Core.dll `
    /reference:System.Drawing.dll `
    /reference:System.Windows.Forms.dll `
    /reference:System.Web.Extensions.dll `
    /out:$OutputPath `
    (Join-Path $PSScriptRoot 'Launcher.cs')

if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $OutputPath)) {
    throw "Build failed with exit code $LASTEXITCODE"
}

Write-Output "BUILT=$OutputPath"
