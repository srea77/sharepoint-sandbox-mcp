param(
    [ValidateSet('Serve', 'Authenticate', 'ResolveSite', 'Doctor')]
    [string]$Mode = 'Serve',
    [string]$SiteId,
    [string]$TenantId,
    [string]$SiteUrl,
    [switch]$UseDeviceCode,
    [string]$BootstrapOutputPath
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
$scopes = @('Sites.ReadWrite.All', 'Sites.Manage.All', 'Files.ReadWrite.All')

function Write-Envelope {
    param([hashtable]$Value)
    $json = $Value | ConvertTo-Json -Depth 100 -Compress
    if ($BootstrapOutputPath) {
        [IO.File]::WriteAllText($BootstrapOutputPath, $json, [Text.UTF8Encoding]::new($false))
        return
    }
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

function Connect-TrustedGraph {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $parameters = @{ Scopes = $scopes; ContextScope = 'CurrentUser'; NoWelcome = $true }
    if ($TenantId) { $parameters.TenantId = $TenantId }
    if ($UseDeviceCode) { $parameters.UseDeviceCode = $true }
    Connect-MgGraph @parameters | Out-Null
}

function Get-SafeRelativePath {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value) -or $Value.Length -gt 2048) { throw 'A non-empty bounded site-relative Graph path is required.' }
    if ($Value.StartsWith('/') -or $Value.StartsWith('\') -or $Value.Contains('://') -or $Value.Contains('\')) { throw 'Absolute, rooted, and backslash Graph paths are rejected.' }
    $decoded = $Value
    for ($i = 0; $i -lt 3; $i++) {
        $next = [Uri]::UnescapeDataString($decoded)
        if ($next -eq $decoded) { break }
        $decoded = $next
    }
    $decoded = $decoded.Replace('\', '/')
    if ($decoded -match '(^|/)\.\.(/|$)' -or $decoded -match '(?i)graph\.microsoft\.com|sharepoint\.(com|de|cn|us)') { throw 'Graph path confinement rejected the request.' }
    return $Value.TrimStart('/')
}

try {
    if ($Mode -eq 'Doctor') {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Write-Envelope @{ ok = $true; moduleVersion = (Get-Module Microsoft.Graph.Authentication).Version.ToString(); context = [bool](Get-MgContext) }
        exit 0
    }

    Connect-TrustedGraph
    $context = Get-MgContext

    if ($Mode -eq 'Authenticate') {
        Write-Envelope @{ ok = $true; tenantId = $context.TenantId; account = $context.Account; authType = $context.AuthType; scopes = @($context.Scopes) }
        exit 0
    }

    if ($Mode -eq 'ResolveSite') {
        $uri = [Uri]$SiteUrl
        if ($uri.Scheme -ne 'https' -or $uri.Host -notmatch '(?i)^[a-z0-9][a-z0-9-]*\.sharepoint\.(com|de|cn|us)$' -or $uri.Query -or $uri.Fragment) { throw 'SiteUrl is not a canonical SharePoint Online HTTPS URL.' }
        $path = if ($uri.AbsolutePath -eq '/') { '/' } else { $uri.AbsolutePath.TrimEnd('/') }
        $result = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($uri.Host):$path" -OutputType PSObject
        Write-Envelope @{ ok = $true; data = $result; tenantId = $context.TenantId; account = $context.Account }
        exit 0
    }

    if ([string]::IsNullOrWhiteSpace($SiteId) -or $SiteId.Contains('/') -or $SiteId.Contains('\')) { throw 'A canonical configured Graph site ID is required.' }
    $escapedSiteId = [Uri]::EscapeDataString($SiteId)
    Write-Envelope @{ ok = $true; ready = $true; tenantId = $context.TenantId; account = $context.Account }

    while (($line = [Console]::In.ReadLine()) -ne $null) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $correlationId = [Guid]::NewGuid().ToString()
        $tempInput = $null
        $tempOutput = $null
        try {
            $request = $line | ConvertFrom-Json -Depth 100
            $method = ([string]$request.method).ToUpperInvariant()
            if ($method -notin @('GET', 'POST', 'PUT', 'PATCH', 'DELETE')) { throw 'Unsupported HTTP method.' }
            $relativePath = Get-SafeRelativePath ([string]$request.relativePath)
            $targetUri = "https://graph.microsoft.com/v1.0/sites/$escapedSiteId/$relativePath"
            $invoke = @{ Method = $method; Uri = $targetUri; OutputType = 'PSObject'; ResponseHeadersVariable = 'responseHeaders'; StatusCodeVariable = 'statusCode' }
            if ($request.expectedETag) {
                $expectedETag = [string]$request.expectedETag
                if ($expectedETag -ne '*' -and ($expectedETag.Length -lt 2 -or $expectedETag.Length -gt 512 -or -not $expectedETag.StartsWith('"') -or -not $expectedETag.EndsWith('"') -or $expectedETag -match '[\r\n]')) { throw 'expectedETag must be * or a quoted HTTP entity tag.' }
                $invoke.Headers = @{ 'If-Match' = $expectedETag }
            }
            if ($request.binaryBase64) {
                $tempInput = Join-Path ([IO.Path]::GetTempPath()) ("sharepoint-mcp-input-" + [Guid]::NewGuid().ToString('N'))
                [IO.File]::WriteAllBytes($tempInput, [Convert]::FromBase64String([string]$request.binaryBase64))
                $invoke.InputFilePath = $tempInput
                $invoke.ContentType = if ($request.contentType) { [string]$request.contentType } else { 'application/octet-stream' }
            } elseif ($null -ne $request.body) {
                $invoke.Body = ($request.body | ConvertTo-Json -Depth 100 -Compress)
                $invoke.ContentType = 'application/json'
            }
            if ($request.expectBinary) {
                $tempOutput = Join-Path ([IO.Path]::GetTempPath()) ("sharepoint-mcp-output-" + [Guid]::NewGuid().ToString('N'))
                $invoke.OutputFilePath = $tempOutput
                $invoke.Remove('OutputType')
            }
            $result = Invoke-MgGraphRequest @invoke
            $requestId = $responseHeaders['request-id'] | Select-Object -First 1
            $responseETag = $responseHeaders['ETag'] | Select-Object -First 1
            if ($request.expectBinary) {
                Write-Envelope @{ ok = $true; statusCode = $statusCode; correlationId = $(if ($requestId) { $requestId } else { $correlationId }); eTag = $responseETag; binaryBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($tempOutput)) }
            } else {
                Write-Envelope @{ ok = $true; statusCode = $statusCode; correlationId = $(if ($requestId) { $requestId } else { $correlationId }); eTag = $responseETag; data = $result }
            }
        } catch {
            Write-Envelope @{ ok = $false; correlationId = $correlationId; error = $_.Exception.Message }
        } finally {
            if ($tempInput -and (Test-Path -LiteralPath $tempInput)) { Remove-Item -LiteralPath $tempInput -Force }
            if ($tempOutput -and (Test-Path -LiteralPath $tempOutput)) { Remove-Item -LiteralPath $tempOutput -Force }
        }
    }
} catch {
    Write-Envelope @{ ok = $false; fatal = $true; error = $_.Exception.Message }
    exit 1
}
