param(
    [string]$Version = '2.90.0'
)

$ErrorActionPreference = 'Stop'
$expectedVersion = '2.90.0'
$expectedSha256 = 'c4ef59b14f0edd074427fd9981e57b0780965ccdcf6191c033fdf4b4361f33d7'
if ($Version -ne $expectedVersion) {
    throw "Only the reviewed Azure CLI version $expectedVersion is supported by this installer."
}

$installRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'Programs\AzureCLI'))
$versionRoot = [IO.Path]::GetFullPath((Join-Path $installRoot $Version))
$azureCli = Join-Path $versionRoot 'bin\az.cmd'
if (-not $versionRoot.StartsWith($installRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to install outside the current user's Azure CLI directory."
}

if (-not (Test-Path -LiteralPath $azureCli)) {
    $temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("sharepoint-mcp-azure-cli-" + [Guid]::NewGuid().ToString('N'))
    $archivePath = Join-Path $temporaryRoot "azure-cli-$Version-x64.zip"
    try {
        New-Item -ItemType Directory -Path $temporaryRoot -Force | Out-Null
        Write-Host "Downloading Microsoft Azure CLI $Version from Microsoft's official distribution..."
        Invoke-WebRequest -Uri "https://azcliprod.blob.core.windows.net/zip/azure-cli-$Version-x64.zip" -OutFile $archivePath -UseBasicParsing
        $actualSha256 = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actualSha256 -ne $expectedSha256) {
            throw "Azure CLI archive integrity check failed. Expected $expectedSha256, received $actualSha256."
        }
        New-Item -ItemType Directory -Path $versionRoot -Force | Out-Null
        Expand-Archive -LiteralPath $archivePath -DestinationPath $versionRoot -Force
    }
    catch {
        if (Test-Path -LiteralPath $versionRoot) { Remove-Item -LiteralPath $versionRoot -Recurse -Force }
        throw
    }
    finally {
        if (Test-Path -LiteralPath $temporaryRoot) { Remove-Item -LiteralPath $temporaryRoot -Recurse -Force }
    }
}

if (-not (Test-Path -LiteralPath $azureCli)) {
    throw "Azure CLI installation did not produce $azureCli"
}

$binPath = Split-Path -Parent $azureCli
$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
if ($null -eq $userPath) { $userPath = '' }
$entries = $userPath.Split([IO.Path]::PathSeparator, [StringSplitOptions]::RemoveEmptyEntries)
if ($entries -notcontains $binPath) {
    [Environment]::SetEnvironmentVariable('Path', (($entries + $binPath) -join [IO.Path]::PathSeparator), 'User')
}

Write-Host "Azure CLI $Version is ready for this Windows user at $azureCli"
Write-Output $azureCli
