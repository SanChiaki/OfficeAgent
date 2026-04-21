[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [string[]]$Architectures = @("x86", "x64")
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Select-MsBuildExe {
    $editions = @("Enterprise", "Professional", "Community", "BuildTools", "TestAgent")
    foreach ($edition in $editions) {
        $path = "C:\Program Files\Microsoft Visual Studio\2022\$edition\MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path $path) { return $path }
        $x86Path = "C:\Program Files (x86)\Microsoft Visual Studio\2022\$edition\MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path $x86Path) { return $x86Path }
    }
    # Fallback: try to find MSBuild.exe via vswhere
    $vswherePath = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswherePath) {
        $installPath = & $vswherePath -latest -property installationPath -products * 2>$null
        if ($installPath) {
            $msbuild = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuild) { return $msbuild }
        }
    }
    # Final fallback: use .NET Framework MSBuild
    $dotnetMsbuild = Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
    if (Test-Path $dotnetMsbuild) { return $dotnetMsbuild }
    throw "Could not find MSBuild. Please ensure Visual Studio is installed."
}

$frontendRoot = Join-Path $repoRoot "src\\OfficeAgent.Frontend"
$addinProject = Join-Path $repoRoot "src\\OfficeAgent.ExcelAddIn\\OfficeAgent.ExcelAddIn.csproj"
$addinOutputRoot = Join-Path $repoRoot "src\\OfficeAgent.ExcelAddIn\\bin\\$Configuration"
$payloadRoot = Join-Path $repoRoot "artifacts\\installer\\payload"
$outputRoot = Join-Path $repoRoot "artifacts\\installer"
$bundleRoot = Join-Path $repoRoot "installer\\OfficeAgent.SetupBundle"
$bundleSource = Join-Path $bundleRoot "Bundle.wxs"
$bundlePrereqRoot = Join-Path $bundleRoot "prereqs"
$offlineSetupPath = Join-Path $outputRoot "OfficeAgent.Setup.exe"
$offlineSetupWixPdbPath = Join-Path $outputRoot "OfficeAgent.Setup.wixpdb"
$toolsManifestPath = Join-Path $repoRoot ".config\\dotnet-tools.json"
$wixSource = Join-Path $PSScriptRoot "Product.wxs"
$msbuild = Select-MsBuildExe
$buildVstoAddInScript = Join-Path $repoRoot "eng\\Build-VstoAddIn.ps1"

function Invoke-NativeCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(ValueFromRemainingArguments = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath $($Arguments -join ' ')"
    }
}

function Ensure-WixExtension {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExtensionReference
    )

    Write-Host "Ensuring WiX extension $ExtensionReference..."
    Invoke-NativeCommand "dotnet" "wix" "extension" "add" "--global" $ExtensionReference
}

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (!(Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Missing ${Description}: $Path"
    }
}

function Build-MsiForArchitecture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Architecture,
        [string]$ProductVersion = "1.0.0"
    )

    $normalizedArchitecture = $Architecture.Trim().ToLowerInvariant()
    if ($normalizedArchitecture -notin @("x86", "x64")) {
        throw "Unsupported architecture '$Architecture'. Expected x86 or x64."
    }

    $msiPath = Join-Path $outputRoot ("OfficeAgent.Setup-{0}.msi" -f $normalizedArchitecture)
    $wixPdbPath = Join-Path $outputRoot ("OfficeAgent.Setup-{0}.wixpdb" -f $normalizedArchitecture)
    if (Test-Path $msiPath) {
        Remove-Item -Force $msiPath
    }

    if (Test-Path $wixPdbPath) {
        Remove-Item -Force $wixPdbPath
    }

    Write-Host "Building MSI for $normalizedArchitecture..."
    Invoke-NativeCommand "dotnet" "wix" "build" $wixSource "-arch" $normalizedArchitecture "-d" "PublishRoot=$payloadRoot" "-d" "ProductVersion=$ProductVersion" "-o" $msiPath
    return $msiPath
}

Write-Host "Using MSBuild: $msbuild"
Assert-FileExists -Path $toolsManifestPath -Description "dotnet tools manifest"
$toolsManifest = Get-Content -Raw $toolsManifestPath | ConvertFrom-Json
$wixToolVersion = $toolsManifest.tools.wix.version
$wixBalExtension = "WixToolset.Bal.wixext"
$wixUtilExtension = "WixToolset.Util.wixext"
$wixExtensionCacheRoot = Join-Path $env:USERPROFILE ".wix\\extensions"
$wixExtensionCacheFolderName = "wixext{0}" -f $wixToolVersion.Split('.')[0]
$wixBalExtensionPath = Join-Path $wixExtensionCacheRoot "$wixBalExtension\\$wixToolVersion\\$wixExtensionCacheFolderName\\WixToolset.BootstrapperApplications.wixext.dll"
$wixUtilExtensionPath = Join-Path $wixExtensionCacheRoot "$wixUtilExtension\\$wixToolVersion\\$wixExtensionCacheFolderName\\WixToolset.Util.wixext.dll"

if ([string]::IsNullOrWhiteSpace($wixToolVersion)) {
    throw "Unable to determine WiX tool version from $toolsManifestPath"
}

$vstoRuntimeInstaller = Join-Path $bundlePrereqRoot "vstor_redist.exe"
$webView2RuntimeInstallerX86 = Join-Path $bundlePrereqRoot "MicrosoftEdgeWebView2RuntimeInstallerX86.exe"
$webView2RuntimeInstallerX64 = Join-Path $bundlePrereqRoot "MicrosoftEdgeWebView2RuntimeInstallerX64.exe"

Assert-FileExists -Path $bundleSource -Description "offline setup bundle source"
Assert-FileExists -Path $vstoRuntimeInstaller -Description "VSTO runtime redistributable"
Assert-FileExists -Path $webView2RuntimeInstallerX86 -Description "WebView2 x86 standalone installer"
Assert-FileExists -Path $webView2RuntimeInstallerX64 -Description "WebView2 x64 standalone installer"

Write-Host "Installing frontend dependencies..."
Push-Location $frontendRoot
try {
    Invoke-NativeCommand "npm.cmd" "install"
}
finally {
    Pop-Location
}

Write-Host "Building frontend..."
Push-Location $frontendRoot
try {
    Invoke-NativeCommand "npm.cmd" "run" "build"
}
finally {
    Pop-Location
}

Write-Host "Restoring installer tools..."
Push-Location $repoRoot
try {
    Invoke-NativeCommand "dotnet" "tool" "restore"
}
finally {
    Pop-Location
}

Write-Host "Ensuring WiX bundle extensions are installed..."
Ensure-WixExtension "$wixBalExtension/$wixToolVersion"
Ensure-WixExtension "$wixUtilExtension/$wixToolVersion"
Assert-FileExists -Path $wixBalExtensionPath -Description "WiX bootstrapper extension"
Assert-FileExists -Path $wixUtilExtensionPath -Description "WiX util extension"

$commitCount = [int](git rev-list --count HEAD).Trim()
$productVersion = "1.0.$commitCount"
Write-Host "App version: $productVersion"

$versionFile = Join-Path $repoRoot "src\\OfficeAgent.ExcelAddIn\\Properties\\Version.g.cs"
$versionContent = @"
using System.Reflection;

[assembly: AssemblyVersion("$productVersion")]
[assembly: AssemblyFileVersion("$productVersion")]

namespace OfficeAgent.ExcelAddIn
{
    internal static class VersionInfo
    {
        public const string AppVersion = "$productVersion";
    }
}
"@
[System.IO.File]::WriteAllText($versionFile, $versionContent)
Write-Host "Generated Version.g.cs with version $productVersion"

Write-Host "Building VSTO add-in..."
Invoke-NativeCommand "pwsh" "-NoProfile" "-ExecutionPolicy" "Bypass" "-File" $buildVstoAddInScript "-ProjectPath" $addinProject "-Configuration" $Configuration "-VisualStudioMSBuildPath" $msbuild

if (!(Test-Path $addinOutputRoot)) {
    throw "Expected add-in output folder not found: $addinOutputRoot"
}

Write-Host "Preparing installer payload..."
if (Test-Path $payloadRoot) {
    Remove-Item -Recurse -Force $payloadRoot
}

New-Item -ItemType Directory -Path $payloadRoot | Out-Null
Copy-Item -Recurse -Force (Join-Path $addinOutputRoot "*") $payloadRoot

$frontendDist = Join-Path $frontendRoot "dist"
$frontendPayload = Join-Path $payloadRoot "frontend"
New-Item -ItemType Directory -Path $frontendPayload | Out-Null
Copy-Item -Recurse -Force (Join-Path $frontendDist "*") $frontendPayload

New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
@(
    (Join-Path $outputRoot "OfficeAgent.Setup.msi"),
    (Join-Path $outputRoot "OfficeAgent.Setup.wixpdb")
) | ForEach-Object {
    if (Test-Path $_) {
        Remove-Item -Force $_
    }
}

Write-Host "Building MSI version $productVersion..."

$builtMsiPaths = @()
foreach ($architecture in $Architectures) {
    $builtMsiPaths += Build-MsiForArchitecture -Architecture $architecture -ProductVersion $productVersion
}

Assert-FileExists -Path (Join-Path $outputRoot "OfficeAgent.Setup-x86.msi") -Description "x86 OfficeAgent MSI"
Assert-FileExists -Path (Join-Path $outputRoot "OfficeAgent.Setup-x64.msi") -Description "x64 OfficeAgent MSI"

if (Test-Path $offlineSetupPath) {
    Remove-Item -Force $offlineSetupPath
}

if (Test-Path $offlineSetupWixPdbPath) {
    Remove-Item -Force $offlineSetupWixPdbPath
}

Write-Host "Building offline setup bundle..."
Invoke-NativeCommand "dotnet" "wix" "build" $bundleSource "-arch" "x86" "-ext" $wixBalExtensionPath "-ext" $wixUtilExtensionPath "-bindpath" $bundleRoot "-bindpath" $outputRoot "-d" "ProductVersion=$productVersion" "-o" $offlineSetupPath

Write-Host "Installer outputs created at:"
$builtMsiPaths | ForEach-Object { Write-Host " - $_" }
Write-Host " - $offlineSetupPath"
