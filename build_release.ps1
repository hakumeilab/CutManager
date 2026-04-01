param(
    [switch]$InstallDependencies
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$versionFile = Join-Path $repoRoot "cutmanager\__init__.py"
$pythonExe = Join-Path $repoRoot ".venv\Scripts\python.exe"
$deployExe = Join-Path $repoRoot ".venv\Scripts\pyside6-deploy.exe"
$distDir = Join-Path $repoRoot "dist"

$versionMatch = Select-String -Path $versionFile -Pattern '__version__ = "(?<version>[^"]+)"'
if (-not $versionMatch) {
    throw "Version was not found in cutmanager\\__init__.py."
}

$version = $versionMatch.Matches[0].Groups["version"].Value
$zipPath = Join-Path $distDir "CutManager-$version-windows-standalone.zip"
$hashPath = Join-Path $distDir "CutManager-$version-windows-standalone.sha256.txt"

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python executable was not found: $pythonExe"
}

if (-not (Test-Path -LiteralPath $deployExe)) {
    throw "pyside6-deploy.exe was not found: $deployExe"
}

Push-Location $repoRoot
try {
    if ($InstallDependencies) {
        & $pythonExe -m pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) {
            throw "Dependency installation failed with exit code $LASTEXITCODE."
        }
    }

    & $deployExe -c pysidedeploy.spec -f
    if ($LASTEXITCODE -ne 0) {
        throw "pyside6-deploy failed with exit code $LASTEXITCODE."
    }

    $buildOutputDir = @(
        (Join-Path $repoRoot "deployment"),
        (Join-Path $repoRoot "CutManager.dist")
    ) |
        Where-Object { Test-Path -LiteralPath $_ } |
        Sort-Object { (Get-Item -LiteralPath $_).LastWriteTimeUtc } -Descending |
        Select-Object -First 1

    if (-not $buildOutputDir) {
        throw "Build succeeded but no output directory was found."
    }

    $expectedExe = Join-Path $buildOutputDir "CutManager.exe"
    $fallbackExe = Join-Path $buildOutputDir "main.exe"
    if (Test-Path -LiteralPath $fallbackExe) {
        if (Test-Path -LiteralPath $expectedExe) {
            Remove-Item -LiteralPath $expectedExe -Force
        }
        Move-Item -LiteralPath $fallbackExe -Destination $expectedExe -Force
    }

    if (-not (Test-Path -LiteralPath $expectedExe)) {
        throw "Build succeeded but CutManager.exe was not found in $buildOutputDir."
    }

    New-Item -ItemType Directory -Path $distDir -Force | Out-Null

    if (Test-Path -LiteralPath $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }

    if (Test-Path -LiteralPath $hashPath) {
        Remove-Item -LiteralPath $hashPath -Force
    }

    Compress-Archive -Path (Join-Path $buildOutputDir "*") -DestinationPath $zipPath -Force

    $hash = Get-FileHash -LiteralPath $zipPath -Algorithm SHA256
    $hashLine = "{0} *{1}" -f $hash.Hash.ToLowerInvariant(), [System.IO.Path]::GetFileName($zipPath)
    Set-Content -LiteralPath $hashPath -Value $hashLine -Encoding ASCII

    Write-Host "Release assets created."
    Write-Host "Version : $version"
    Write-Host "Build   : $buildOutputDir"
    Write-Host "ZIP     : $zipPath"
    Write-Host "SHA256  : $hashPath"
}
finally {
    Pop-Location
}
