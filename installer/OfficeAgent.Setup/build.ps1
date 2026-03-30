[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [string[]]$Architectures = @("x86", "x64")
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$frontendRoot = Join-Path $repoRoot "src\\OfficeAgent.Frontend"
$addinProject = Join-Path $repoRoot "src\\OfficeAgent.ExcelAddIn\\OfficeAgent.ExcelAddIn.csproj"
$addinOutputRoot = Join-Path $repoRoot "src\\OfficeAgent.ExcelAddIn\\bin\\$Configuration"
$payloadRoot = Join-Path $repoRoot "artifacts\\installer\\payload"
$outputRoot = Join-Path $repoRoot "artifacts\\installer"
$wixSource = Join-Path $PSScriptRoot "Product.wxs"
$msbuild = "C:\\Program Files\\Microsoft Visual Studio\\2022\\Community\\MSBuild\\Current\\Bin\\MSBuild.exe"

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

Write-Host "Building frontend..."
Push-Location $frontendRoot
try {
    Invoke-NativeCommand "npm.cmd" "run" "build"
}
finally {
    Pop-Location
}

Write-Host "Building VSTO add-in..."
Invoke-NativeCommand $msbuild $addinProject "/restore" "/p:RestorePackagesConfig=true" "/p:Configuration=$Configuration"

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
Write-Host "Building MSI version $productVersion..."

$builtMsiPaths = @()
foreach ($architecture in $Architectures) {
    $builtMsiPaths += Build-MsiForArchitecture -Architecture $architecture -ProductVersion $productVersion
}

Write-Host "MSI created at:"
$builtMsiPaths | ForEach-Object { Write-Host " - $_" }
