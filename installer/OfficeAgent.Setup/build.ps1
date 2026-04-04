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
$wixSource = Join-Path $PSScriptRoot "Product.wxs"
$msbuild = Select-MsBuildExe

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

$manifestThumbprint = $null
$certificateCreated = $false

try {
    $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My","CurrentUser")
    $certStore.Open("ReadOnly")
    $existingCert = $certStore.Certificates | Where-Object {
        $_.EnhancedKeyUsageList | ForEach-Object { $_.Value } | Where-Object { $_ -eq "1.3.6.1.5.5.7.3.3" } |
        Where-Object { $_.Subject -match "OfficeAgent" }
    } | Sort-Object NotBefore -Descending | Select-Object -First 1
    $certStore.Close()

    if ($existingCert) {
        $manifestThumbprint = $existingCert.Thumbprint
        Write-Host "Using existing certificate: $manifestThumbprint"
    } else {
        Write-Host "Generating temporary code signing certificate..."
        $cert = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject "CN=OfficeAgent" `
            -CertStoreLocation "Cert:\CurrentUser\My"
        $manifestThumbprint = $cert.Thumbprint
        $certificateCreated = $true
        Write-Host "Certificate thumbprint: $manifestThumbprint"
    }
} catch {
    throw "Failed to get or create code signing certificate: $_"
}

try {
    Write-Host "Restoring installer tools..."
    Push-Location $repoRoot
    try {
        Invoke-NativeCommand "dotnet" "tool" "restore"
    }
    finally {
        Pop-Location
    }

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
    $msbuildArgs = @(
        $addinProject
        "/restore"
        "/p:RestorePackagesConfig=true"
        "/p:Configuration=$Configuration"
        "/p:ManifestCertificateThumbprint=$manifestThumbprint"
    )
    Invoke-NativeCommand $msbuild @msbuildArgs

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

    Write-Host "MSI created at:"
    $builtMsiPaths | ForEach-Object { Write-Host " - $_" }
} finally {
    if ($certificateCreated -and $manifestThumbprint) {
        Write-Host "Removing temporary certificate..."
        Remove-Item "Cert:\CurrentUser\My\$manifestThumbprint" -ErrorAction SilentlyContinue
    }
}
