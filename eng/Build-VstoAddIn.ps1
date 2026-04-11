[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath,

    [string]$Configuration = "Debug",

    [string]$VisualStudioMSBuildPath
)

$ErrorActionPreference = "Stop"

function Select-MsBuildExe {
    $editions = @("Enterprise", "Professional", "Community", "BuildTools", "TestAgent")
    foreach ($edition in $editions) {
        $path = "C:\Program Files\Microsoft Visual Studio\2022\$edition\MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path $path) { return $path }

        $x86Path = "C:\Program Files (x86)\Microsoft Visual Studio\2022\$edition\MSBuild\Current\Bin\MSBuild.exe"
        if (Test-Path $x86Path) { return $x86Path }
    }

    $vswherePath = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswherePath) {
        $installPath = & $vswherePath -latest -property installationPath -products * 2>$null
        if ($installPath) {
            $msbuild = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuild) { return $msbuild }
        }
    }

    $dotnetMsbuild = Join-Path $env:SystemRoot "Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
    if (Test-Path $dotnetMsbuild) { return $dotnetMsbuild }

    throw "Could not find MSBuild. Please ensure Visual Studio is installed."
}

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

$resolvedProjectPath = [System.IO.Path]::GetFullPath($ProjectPath)
if (!(Test-Path $resolvedProjectPath)) {
    throw "Project file was not found: $resolvedProjectPath"
}

if ([string]::IsNullOrWhiteSpace($VisualStudioMSBuildPath)) {
    $VisualStudioMSBuildPath = Select-MsBuildExe
}

Write-Host "Using MSBuild: $VisualStudioMSBuildPath"
Write-Host "Building signed VSTO project: $resolvedProjectPath"

$manifestThumbprint = $null

try {
    $subject = "CN=OfficeAgent Temporary VSTO Build $([Guid]::NewGuid().ToString('N'))"
    $certificate = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $subject `
        -CertStoreLocation "Cert:\CurrentUser\My"
    $manifestThumbprint = $certificate.Thumbprint

    Write-Host "Generated temporary manifest certificate: $manifestThumbprint"

    $msbuildArgs = @(
        $resolvedProjectPath
        "/restore"
        "/p:RestorePackagesConfig=true"
        "/p:Configuration=$Configuration"
        "/p:ManifestCertificateThumbprint=$manifestThumbprint"
    )

    Invoke-NativeCommand $VisualStudioMSBuildPath @msbuildArgs
}
finally {
    if ($manifestThumbprint) {
        Write-Host "Removing temporary manifest certificate..."
        Remove-Item "Cert:\CurrentUser\My\$manifestThumbprint" -ErrorAction SilentlyContinue
    }
}
