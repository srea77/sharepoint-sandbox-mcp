param(
    [string]$Runtime = 'win-x64',
    [string]$Output = (Join-Path $PSScriptRoot '..\artifacts\publish\win-x64')
)

$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot '..\src\SharePointMcp.Server\SharePointMcp.Server.csproj'
dotnet publish $project --configuration Release --runtime $Runtime --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true --output $Output
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed with exit code $LASTEXITCODE" }
Write-Host "Published SharePoint MCP to $([IO.Path]::GetFullPath($Output))"
