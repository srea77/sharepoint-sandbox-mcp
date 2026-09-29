param(
    [string]$Runtime = 'win-x64',
    [string]$Version = '1.2.0',
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
$mcpbCliVersion = '2.1.2'
$nodeVersion = '22.22.3'
$nodeArchiveSha256 = '6c8d54f635feff4df76c2ca80f45332eb2ff57d25226edce36592e51a177ee33'
$flowAgentVersion = '3.1.0'
$flowAgentSha256 = '85b69e290528c3d3f0d410eaf91c2dcddec64809277a5c54076797d35cc18c47'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$artifactsParent = [IO.Path]::GetFullPath((Join-Path $repoRoot 'artifacts'))
$outputRoot = [IO.Path]::GetFullPath((Join-Path $artifactsParent 'cowork'))
$mcpbStageRoot = Join-Path $outputRoot 'mcpb-unpacked'
$pluginStageRoot = Join-Path $outputRoot 'plugin-unpacked'
$serverRoot = Join-Path $mcpbStageRoot 'server'
$pluginSourceRoot = Join-Path $repoRoot 'plugin'
$kitName = "SharePoint-Sandbox-Cowork-$Version-$Runtime"
$kitRoot = Join-Path $outputRoot $kitName
$mcpbName = "SharePoint-Sandbox-$Version-$Runtime.mcpb"
$mcpbPath = Join-Path $outputRoot $mcpbName
$pluginName = 'sharepoint-sandbox.plugin'
$pluginPath = Join-Path $outputRoot $pluginName
$zipPath = Join-Path $outputRoot "$kitName.zip"
$projectPath = Join-Path $repoRoot 'src\SharePointMcp.Server\SharePointMcp.Server.csproj'
$mcpbManifestSource = Join-Path $repoRoot 'packaging\cowork\manifest.json'
$pluginManifestSource = Join-Path $pluginSourceRoot '.claude-plugin\plugin.json'

if (-not $outputRoot.StartsWith($artifactsParent + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to replace an output directory outside the repository artifacts directory: $outputRoot"
}

$mcpbVersion = (Get-Content -LiteralPath $mcpbManifestSource -Raw | ConvertFrom-Json).version
$pluginVersion = (Get-Content -LiteralPath $pluginManifestSource -Raw | ConvertFrom-Json).version
if ($Version -ne $mcpbVersion -or $Version -ne $pluginVersion) {
    throw "Version $Version must match both the Cowork plugin manifest ($pluginVersion) and MCPB compatibility manifest ($mcpbVersion)."
}

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw 'The .NET 9 SDK is required to build the plugin. End users do not need it.'
}
if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
    throw 'The Claude Code CLI is required only to validate the native plugin structure. End users install through Cowork.'
}
if (-not (Get-Command npx.cmd -ErrorAction SilentlyContinue)) {
    throw 'Node.js/npm is required only to build the optional MCPB compatibility artifact. End users do not need it.'
}

if (Test-Path -LiteralPath $outputRoot) {
    Remove-Item -LiteralPath $outputRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $serverRoot -Force | Out-Null
New-Item -ItemType Directory -Path $pluginStageRoot -Force | Out-Null
New-Item -ItemType Directory -Path $kitRoot -Force | Out-Null

$cacheRoot = Join-Path $artifactsParent 'cache'
$nodeArchive = Join-Path $cacheRoot "node-v$nodeVersion-win-x64.zip"
New-Item -ItemType Directory -Path $cacheRoot -Force | Out-Null
if (-not (Test-Path -LiteralPath $nodeArchive) -or
    (Get-FileHash -LiteralPath $nodeArchive -Algorithm SHA256).Hash.ToLowerInvariant() -ne $nodeArchiveSha256) {
    Write-Host "Downloading pinned Node.js $nodeVersion runtime..."
    Invoke-WebRequest -Uri "https://nodejs.org/dist/v$nodeVersion/node-v$nodeVersion-win-x64.zip" -OutFile $nodeArchive -UseBasicParsing
}
$actualNodeHash = (Get-FileHash -LiteralPath $nodeArchive -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualNodeHash -ne $nodeArchiveSha256) {
    throw "Pinned Node.js archive integrity check failed. Expected $nodeArchiveSha256, received $actualNodeHash."
}

Write-Host "Publishing a self-contained $Runtime server..."
& dotnet publish $projectPath --configuration Release --runtime $Runtime --self-contained true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:DebugType=None -p:DebugSymbols=false --output $serverRoot
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed with exit code $LASTEXITCODE" }

$serverExe = Join-Path $serverRoot 'sharepoint-mcp.exe'
if (-not (Test-Path -LiteralPath $serverExe)) {
    throw "The self-contained server executable was not produced: $serverExe"
}

$flowAgentSourceRoot = Join-Path $repoRoot "third_party\microsoft-power-automate-flowagent\$flowAgentVersion"
$flowAgentSource = Join-Path $flowAgentSourceRoot 'mcp.mjs'
if (-not (Test-Path -LiteralPath $flowAgentSource)) {
    throw "The reviewed Microsoft FlowAgent bundle is missing: $flowAgentSource"
}
$actualFlowAgentHash = (Get-FileHash -LiteralPath $flowAgentSource -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualFlowAgentHash -ne $flowAgentSha256) {
    throw "Pinned FlowAgent integrity check failed. Expected $flowAgentSha256, received $actualFlowAgentHash."
}
$bundledFlowAgentRoot = Join-Path $serverRoot 'flowagent'
New-Item -ItemType Directory -Path $bundledFlowAgentRoot -Force | Out-Null
Copy-Item -LiteralPath $flowAgentSource -Destination $bundledFlowAgentRoot
Copy-Item -LiteralPath (Join-Path $flowAgentSourceRoot 'LICENSE') -Destination (Join-Path $bundledFlowAgentRoot 'LICENSE.microsoft-power-platform-skills.txt')
Copy-Item -LiteralPath (Join-Path $flowAgentSourceRoot 'UPSTREAM.md') -Destination $bundledFlowAgentRoot

$nodeStage = Join-Path $outputRoot 'node-stage'
Expand-Archive -LiteralPath $nodeArchive -DestinationPath $nodeStage -Force
$nodeSourceRoot = Join-Path $nodeStage "node-v$nodeVersion-win-x64"
$bundledNodeRoot = Join-Path $bundledFlowAgentRoot 'node'
New-Item -ItemType Directory -Path $bundledNodeRoot -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $nodeSourceRoot 'node.exe') -Destination $bundledNodeRoot
Copy-Item -LiteralPath (Join-Path $nodeSourceRoot 'LICENSE') -Destination (Join-Path $bundledNodeRoot 'LICENSE.node.txt')
Remove-Item -LiteralPath $nodeStage -Recurse -Force

Copy-Item -LiteralPath $mcpbManifestSource -Destination (Join-Path $mcpbStageRoot 'manifest.json')
Copy-Item -LiteralPath (Join-Path $repoRoot 'packaging\cowork\USER_GUIDE.md') -Destination $mcpbStageRoot

Copy-Item -LiteralPath (Join-Path $pluginSourceRoot '.claude-plugin') -Destination $pluginStageRoot -Recurse
Copy-Item -LiteralPath (Join-Path $pluginSourceRoot '.mcp.json') -Destination $pluginStageRoot
Copy-Item -LiteralPath (Join-Path $pluginSourceRoot 'skills') -Destination $pluginStageRoot -Recurse
Copy-Item -LiteralPath (Join-Path $pluginSourceRoot 'README.md') -Destination $pluginStageRoot
Copy-Item -LiteralPath $serverRoot -Destination (Join-Path $pluginStageRoot 'server') -Recurse

Write-Host 'Verifying the packaged executable launches without a machine-wide .NET installation...'
& $serverExe --help | Out-Host
if ($LASTEXITCODE -ne 0) { throw "The packaged executable smoke test failed with exit code $LASTEXITCODE" }

if (-not $SkipTests) {
    Write-Host 'Running automated tests before packaging...'
    & dotnet test (Join-Path $repoRoot 'SharePointMcp.sln') --configuration Release
    if ($LASTEXITCODE -ne 0) { throw "dotnet test failed with exit code $LASTEXITCODE" }
}

Write-Host 'Validating the native Claude/Cowork plugin structure...'
& claude plugin validate $pluginStageRoot
if ($LASTEXITCODE -ne 0) { throw "Claude plugin validation failed with exit code $LASTEXITCODE" }

Write-Host 'Packing the native Cowork plugin with skills and MCP connector...'
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::CreateFromDirectory(
    $pluginStageRoot,
    $pluginPath,
    [IO.Compression.CompressionLevel]::Optimal,
    $false)

$pluginArchive = [IO.Compression.ZipFile]::OpenRead($pluginPath)
try {
    $pluginEntries = @($pluginArchive.Entries | ForEach-Object { $_.FullName.Replace('\', '/') })
    foreach ($requiredEntry in @(
        '.claude-plugin/plugin.json',
        '.mcp.json',
        'server/sharepoint-mcp.exe',
        'server/flowagent/mcp.mjs',
        'server/flowagent/node/node.exe',
        'server/flowagent/LICENSE.microsoft-power-platform-skills.txt',
        'skills/build-sharepoint-site/SKILL.md',
        'skills/audit-sharepoint-site/SKILL.md',
        'skills/operate-sharepoint-lifecycle/SKILL.md')) {
        if ($pluginEntries -notcontains $requiredEntry) {
            throw "The native plugin archive is missing required entry '$requiredEntry'."
        }
    }
}
finally {
    $pluginArchive.Dispose()
}

Write-Host "Validating the MCPB 0.4 compatibility manifest with @anthropic-ai/mcpb@$mcpbCliVersion..."
& npx.cmd --yes "@anthropic-ai/mcpb@$mcpbCliVersion" validate (Join-Path $mcpbStageRoot 'manifest.json')
if ($LASTEXITCODE -ne 0) { throw "MCPB manifest validation failed with exit code $LASTEXITCODE" }

Write-Host 'Packing the connector-only MCPB compatibility artifact...'
& npx.cmd --yes "@anthropic-ai/mcpb@$mcpbCliVersion" pack $mcpbStageRoot $mcpbPath
if ($LASTEXITCODE -ne 0) { throw "MCPB packing failed with exit code $LASTEXITCODE" }

& npx.cmd --yes "@anthropic-ai/mcpb@$mcpbCliVersion" info $mcpbPath | Out-Host
if ($LASTEXITCODE -ne 0) { throw "MCPB inspection failed with exit code $LASTEXITCODE" }

Copy-Item -LiteralPath $pluginPath -Destination (Join-Path $kitRoot $pluginName)
Copy-Item -LiteralPath $mcpbPath -Destination (Join-Path $kitRoot $mcpbName)
Copy-Item -LiteralPath (Join-Path $repoRoot 'packaging\cowork\setup-sharepoint-sandbox.cmd') -Destination $kitRoot
Copy-Item -LiteralPath (Join-Path $repoRoot 'packaging\cowork\doctor-sharepoint-sandbox.cmd') -Destination $kitRoot
Copy-Item -LiteralPath (Join-Path $repoRoot 'packaging\cowork\setup-power-automate.cmd') -Destination $kitRoot
Copy-Item -LiteralPath (Join-Path $repoRoot 'scripts\install-azure-cli-user.ps1') -Destination $kitRoot
Copy-Item -LiteralPath (Join-Path $repoRoot 'packaging\cowork\DISTRIBUTION_README.txt') -Destination (Join-Path $kitRoot 'README.txt')
Copy-Item -LiteralPath $serverRoot -Destination (Join-Path $kitRoot 'runtime') -Recurse

Compress-Archive -Path (Join-Path $kitRoot '*') -DestinationPath $zipPath -CompressionLevel Optimal

$pluginHash = (Get-FileHash -LiteralPath $pluginPath -Algorithm SHA256).Hash.ToLowerInvariant()
$mcpbHash = (Get-FileHash -LiteralPath $mcpbPath -Algorithm SHA256).Hash.ToLowerInvariant()
$zipHash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
$checksumPath = Join-Path $outputRoot 'SHA256SUMS.txt'
@(
    "$pluginHash *$pluginName"
    "$mcpbHash *$mcpbName"
    "$zipHash *$([IO.Path]::GetFileName($zipPath))"
) | Set-Content -LiteralPath $checksumPath -Encoding Ascii

Write-Host ''
Write-Host 'Cowork plugin built and verified:'
Write-Host "  Native plugin: $pluginPath"
Write-Host "  End-user kit: $zipPath"
Write-Host "  Unpacked plugin: $pluginStageRoot"
Write-Host "  MCPB compatibility artifact: $mcpbPath"
Write-Host "  Checksums: $checksumPath"
