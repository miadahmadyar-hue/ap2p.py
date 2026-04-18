#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InstallDir = "$env:LOCALAPPDATA\ap2p",
    [string]$Ref = "main"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Repo = "miadahmadyar-hue/ap2p.py"
$Archive = "https://github.com/$Repo/archive/$Ref.zip"

function Test-Command($Name) {
    $null -ne (Get-Command $Name -ErrorAction SilentlyContinue)
}

function Write-Info($Message) {
    Write-Host "==> $Message" -ForegroundColor Cyan
}

Write-Info "Installing ap2p.py to $InstallDir"

if (-not (Test-Command "python")) {
    throw "Python is required but was not found on PATH. Install Python 3.8+ from https://www.python.org/ and retry."
}

$PythonVersion = (python -c "import sys; print('{0}.{1}'.format(sys.version_info[0], sys.version_info[1]))").Trim()
Write-Info "Detected Python $PythonVersion"

if (Test-Path $InstallDir) {
    Write-Info "Removing previous install at $InstallDir"
    Remove-Item -Recurse -Force $InstallDir
}
New-Item -ItemType Directory -Path $InstallDir | Out-Null

$TempZip = Join-Path ([System.IO.Path]::GetTempPath()) ("ap2p-" + [System.Guid]::NewGuid().ToString() + ".zip")
$TempExtract = Join-Path ([System.IO.Path]::GetTempPath()) ("ap2p-" + [System.Guid]::NewGuid().ToString())

try {
    Write-Info "Downloading $Archive"
    Invoke-WebRequest -Uri $Archive -OutFile $TempZip -UseBasicParsing

    Write-Info "Extracting archive"
    Expand-Archive -Path $TempZip -DestinationPath $TempExtract -Force

    $SourceRoot = Get-ChildItem -Path $TempExtract -Directory | Select-Object -First 1
    if (-not $SourceRoot) {
        throw "Extracted archive did not contain expected directory."
    }

    Copy-Item -Path (Join-Path $SourceRoot.FullName "*") -Destination $InstallDir -Recurse -Force
}
finally {
    if (Test-Path $TempZip) { Remove-Item -Force $TempZip }
    if (Test-Path $TempExtract) { Remove-Item -Recurse -Force $TempExtract }
}

$Requirements = Join-Path $InstallDir "requirements.txt"
if (Test-Path $Requirements) {
    Write-Info "Installing Python dependencies"
    python -m pip install --user --upgrade -r $Requirements
}

$LauncherDir = Join-Path $InstallDir "bin"
if (-not (Test-Path $LauncherDir)) {
    New-Item -ItemType Directory -Path $LauncherDir | Out-Null
}

$LauncherPath = Join-Path $LauncherDir "ap2p.cmd"
$AppPath = Join-Path $InstallDir "app.py"
$LauncherContent = "@echo off`r`npython `"$AppPath`" %*`r`n"
Set-Content -Path $LauncherPath -Value $LauncherContent -Encoding ASCII

$UserPath = [Environment]::GetEnvironmentVariable("Path", "User")
if (-not ($UserPath -split ";" | Where-Object { $_ -eq $LauncherDir })) {
    Write-Info "Adding $LauncherDir to user PATH"
    $NewPath = if ([string]::IsNullOrEmpty($UserPath)) { $LauncherDir } else { "$UserPath;$LauncherDir" }
    [Environment]::SetEnvironmentVariable("Path", $NewPath, "User")
    $env:Path = "$env:Path;$LauncherDir"
}

Write-Info "ap2p.py installed successfully."
Write-Host ""
Write-Host "Run 'ap2p --help' in a new terminal to get started." -ForegroundColor Green
