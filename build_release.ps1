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
$releaseExePath = Join-Path $distDir "CutManager-$version-windows-onefile.exe"
$hashPath = Join-Path $distDir "CutManager-$version-windows-onefile.sha256.txt"

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

    $candidateExePaths = @(
        (Join-Path $repoRoot "CutManager.exe"),
        (Join-Path $repoRoot "main.exe"),
        (Join-Path $repoRoot "deployment\CutManager.exe"),
        (Join-Path $repoRoot "deployment\main.exe"),
        (Join-Path $repoRoot "CutManager.dist\CutManager.exe"),
        (Join-Path $repoRoot "CutManager.dist\main.exe")
    ) | Where-Object { Test-Path -LiteralPath $_ }

    if (-not $candidateExePaths) {
        throw "Build succeeded but no executable output was found."
    }

    $builtExe = $candidateExePaths |
        Sort-Object { (Get-Item -LiteralPath $_).LastWriteTimeUtc } -Descending |
        Select-Object -First 1

    $builtExeDirectory = Split-Path -Path $builtExe -Parent
    $expectedExe = Join-Path $builtExeDirectory "CutManager.exe"
    if ([System.IO.Path]::GetFileName($builtExe) -ieq "main.exe") {
        if (Test-Path -LiteralPath $expectedExe) {
            Remove-Item -LiteralPath $expectedExe -Force
        }
        Move-Item -LiteralPath $builtExe -Destination $expectedExe -Force
        $builtExe = $expectedExe
    }

    if (-not (Test-Path -LiteralPath $builtExe)) {
        throw "Build succeeded but CutManager.exe was not found."
    }

    New-Item -ItemType Directory -Path $distDir -Force | Out-Null

    if (Test-Path -LiteralPath $releaseExePath) {
        Remove-Item -LiteralPath $releaseExePath -Force
    }

    if (Test-Path -LiteralPath $hashPath) {
        Remove-Item -LiteralPath $hashPath -Force
    }

    Copy-Item -LiteralPath $builtExe -Destination $releaseExePath -Force

    $hash = Get-FileHash -LiteralPath $releaseExePath -Algorithm SHA256
    $hashLine = "{0} *{1}" -f $hash.Hash.ToLowerInvariant(), [System.IO.Path]::GetFileName($releaseExePath)
    Set-Content -LiteralPath $hashPath -Value $hashLine -Encoding ASCII

    Write-Host "Release assets created."
    Write-Host "Version : $version"
    Write-Host "Build   : $builtExeDirectory"
    Write-Host "EXE     : $releaseExePath"
    Write-Host "SHA256  : $hashPath"
}
finally {
    Pop-Location
}
