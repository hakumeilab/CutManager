param(
    [switch]$InstallDependencies
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$versionFile = Join-Path $repoRoot "cutmanager\__init__.py"
$specTemplatePath = Join-Path $repoRoot "pysidedeploy.spec"
$pythonExe = Join-Path $repoRoot ".venv\Scripts\python.exe"
$deployExe = Join-Path $repoRoot ".venv\Scripts\pyside6-deploy.exe"
$distDir = Join-Path $repoRoot "dist"
$defaultIconPath = Join-Path $repoRoot ".venv\Lib\site-packages\PySide6\scripts\deploy_lib\pyside_icon.ico"

$versionMatch = Select-String -Path $versionFile -Pattern '__version__ = "(?<version>[^"]+)"'
if (-not $versionMatch) {
    throw "Version was not found in cutmanager\\__init__.py."
}

$version = $versionMatch.Matches[0].Groups["version"].Value
$releaseOnefileExePath = Join-Path $distDir "CutManager-$version-windows-onefile.exe"
$releaseOnefileHashPath = Join-Path $distDir "CutManager-$version-windows-onefile.sha256.txt"
$releaseSetupExePath = Join-Path $distDir "CutManager-$version-windows-setup.exe"
$releaseSetupHashPath = Join-Path $distDir "CutManager-$version-windows-setup.sha256.txt"

if (-not (Test-Path -LiteralPath $pythonExe)) {
    throw "Python executable was not found: $pythonExe"
}

if (-not (Test-Path -LiteralPath $deployExe)) {
    throw "pyside6-deploy.exe was not found: $deployExe"
}

if (-not (Test-Path -LiteralPath $specTemplatePath)) {
    throw "Deploy spec was not found: $specTemplatePath"
}

function Set-SpecValue {
    param(
        [string[]]$Lines,
        [string]$Section,
        [string]$Key,
        [string]$Value
    )

    $inSection = $false
    for ($lineIndex = 0; $lineIndex -lt $Lines.Count; $lineIndex++) {
        $trimmedLine = $Lines[$lineIndex].Trim()
        if ($trimmedLine -match '^\[(?<name>[^\]]+)\]$') {
            $inSection = $Matches["name"] -eq $Section
            continue
        }

        if ($inSection -and $trimmedLine -match "^$([regex]::Escape($Key))\s*=") {
            $Lines[$lineIndex] = "{0} = {1}" -f $Key, $Value
            return $Lines
        }
    }

    throw "Key '$Key' in section '$Section' was not found in $specTemplatePath."
}

function Get-InnoSetupCompilerPath {
    $command = Get-Command -Name "ISCC.exe" -ErrorAction SilentlyContinue
    if ($command -and $command.Source -and (Test-Path -LiteralPath $command.Source)) {
        return $command.Source
    }

    $candidatePaths = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 5\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 5\ISCC.exe"
    )

    foreach ($candidatePath in $candidatePaths) {
        if ($candidatePath -and (Test-Path -LiteralPath $candidatePath)) {
            return $candidatePath
        }
    }

    return ""
}

function Convert-InnoSetupString {
    param(
        [string]$Value
    )

    return ($Value -replace '"', '""')
}

function Write-ReleaseHashFile {
    param(
        [string]$AssetPath,
        [string]$HashPath
    )

    if (Test-Path -LiteralPath $HashPath) {
        Remove-Item -LiteralPath $HashPath -Force
    }

    $hash = Get-FileHash -LiteralPath $AssetPath -Algorithm SHA256
    $hashLine = "{0} *{1}" -f $hash.Hash.ToLowerInvariant(), [System.IO.Path]::GetFileName($AssetPath)
    Set-Content -LiteralPath $HashPath -Value $hashLine -Encoding ASCII
}

function New-InstallerScriptContent {
    param(
        [string]$Version,
        [string]$SourceExePath,
        [string]$OutputDirectory,
        [string]$IconPath
    )

    $escapedSourceExePath = Convert-InnoSetupString -Value $SourceExePath
    $escapedOutputDirectory = Convert-InnoSetupString -Value $OutputDirectory
    $escapedIconPath = Convert-InnoSetupString -Value $IconPath
    $setupIconLine = if ($IconPath) { "SetupIconFile=$escapedIconPath" } else { "" }

    return @"
[Setup]
AppId={{9F3E0B72-8E7D-4A3D-BC0D-6E0C6415D742}
AppName=CutManager
AppVersion=$Version
AppPublisher=hakumeilab
DefaultDirName={localappdata}\Programs\CutManager
DefaultGroupName=CutManager
DisableProgramGroupPage=yes
OutputDir=$escapedOutputDirectory
OutputBaseFilename=CutManager-$Version-windows-setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\CutManager.exe
$setupIconLine

[Files]
Source: "$escapedSourceExePath"; DestDir: "{app}"; DestName: "CutManager.exe"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\CutManager"; Filename: "{app}\CutManager.exe"
Name: "{autodesktop}\CutManager"; Filename: "{app}\CutManager.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\CutManager.exe"; Description: "Launch CutManager"; Flags: nowait postinstall skipifsilent
"@
}

Push-Location $repoRoot
try {
    if ($InstallDependencies) {
        & $pythonExe -m pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) {
            throw "Dependency installation failed with exit code $LASTEXITCODE."
        }
    }

    $tempSpecPath = Join-Path $env:TEMP "CutManager-release-$version.spec"
    $tempInstallerScriptPath = Join-Path $env:TEMP "CutManager-installer-$version.iss"
    $specContent = Get-Content -LiteralPath $specTemplatePath
    $normalizedIconPath = if (Test-Path -LiteralPath $defaultIconPath) { $defaultIconPath } else { "" }
    $specContent = Set-SpecValue -Lines $specContent -Section "app" -Key "title" -Value "CutManager"
    $specContent = Set-SpecValue -Lines $specContent -Section "app" -Key "project_dir" -Value "."
    $specContent = Set-SpecValue -Lines $specContent -Section "app" -Key "input_file" -Value "main.py"
    $specContent = Set-SpecValue -Lines $specContent -Section "app" -Key "exec_directory" -Value "."
    $specContent = Set-SpecValue -Lines $specContent -Section "app" -Key "icon" -Value $normalizedIconPath
    $specContent = Set-SpecValue -Lines $specContent -Section "python" -Key "python_path" -Value $pythonExe
    $specContent = Set-SpecValue -Lines $specContent -Section "nuitka" -Key "mode" -Value "onefile"
    Set-Content -LiteralPath $tempSpecPath -Value $specContent -Encoding ASCII

    $candidateCleanupPaths = @(
        (Join-Path $repoRoot "CutManager.exe"),
        (Join-Path $repoRoot "main.exe"),
        (Join-Path $repoRoot "deployment\CutManager.exe"),
        (Join-Path $repoRoot "deployment\main.exe"),
        (Join-Path $repoRoot "CutManager.dist\CutManager.exe"),
        (Join-Path $repoRoot "CutManager.dist\main.exe"),
        (Join-Path $repoRoot "main.dist\main.exe")
    )

    foreach ($cleanupPath in $candidateCleanupPaths) {
        if (Test-Path -LiteralPath $cleanupPath) {
            Remove-Item -LiteralPath $cleanupPath -Force
        }
    }

    & $deployExe -c $tempSpecPath -f
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

    if (Test-Path -LiteralPath $releaseOnefileExePath) {
        Remove-Item -LiteralPath $releaseOnefileExePath -Force
    }

    if (Test-Path -LiteralPath $releaseSetupExePath) {
        Remove-Item -LiteralPath $releaseSetupExePath -Force
    }

    Copy-Item -LiteralPath $builtExe -Destination $releaseOnefileExePath -Force
    Write-ReleaseHashFile -AssetPath $releaseOnefileExePath -HashPath $releaseOnefileHashPath

    $innoSetupCompilerPath = Get-InnoSetupCompilerPath
    if ($innoSetupCompilerPath) {
        $installerScriptContent = New-InstallerScriptContent `
            -Version $version `
            -SourceExePath $releaseOnefileExePath `
            -OutputDirectory $distDir `
            -IconPath $normalizedIconPath
        Set-Content -LiteralPath $tempInstallerScriptPath -Value $installerScriptContent -Encoding ASCII

        & $innoSetupCompilerPath $tempInstallerScriptPath
        if ($LASTEXITCODE -ne 0) {
            throw "Inno Setup build failed with exit code $LASTEXITCODE."
        }

        if (-not (Test-Path -LiteralPath $releaseSetupExePath)) {
            throw "Installer build succeeded but setup executable was not found: $releaseSetupExePath"
        }

        Write-ReleaseHashFile -AssetPath $releaseSetupExePath -HashPath $releaseSetupHashPath
    }
    else {
        Write-Warning "ISCC.exe was not found. Skipped setup installer build."
    }

    Write-Host "Release assets created."
    Write-Host "Version : $version"
    Write-Host "Build   : $builtExeDirectory"
    Write-Host "Onefile : $releaseOnefileExePath"
    Write-Host "SHA256  : $releaseOnefileHashPath"
    if (Test-Path -LiteralPath $releaseSetupExePath) {
        Write-Host "Setup   : $releaseSetupExePath"
        Write-Host "SHA256  : $releaseSetupHashPath"
    }
}
finally {
    if ($tempInstallerScriptPath -and (Test-Path -LiteralPath $tempInstallerScriptPath)) {
        Remove-Item -LiteralPath $tempInstallerScriptPath -Force
    }
    if ($tempSpecPath -and (Test-Path -LiteralPath $tempSpecPath)) {
        Remove-Item -LiteralPath $tempSpecPath -Force
    }
    Pop-Location
}
