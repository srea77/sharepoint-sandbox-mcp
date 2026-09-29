using System.Diagnostics;
using System.Text.Json.Nodes;
using Azure.Core;
using Azure.Identity;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

/// <summary>
/// Azure CLI token bridge used only inside the trusted process. It accepts the
/// Graph audience or one exact configured SharePoint hostname and never emits
/// the returned token to logs, configuration, or MCP results.
/// </summary>
internal sealed class BoundAzureCliCredential : TokenCredential
{
    private readonly string? _tenantId;
    private readonly string? _sharePointHostname;
    private readonly string _azureCli;

    public BoundAzureCliCredential(string tenantId, string[] scopes)
    {
        _tenantId = Guid.TryParse(tenantId, out var parsed) ? parsed.ToString() : null;
        _sharePointHostname = RequireResource(scopes).Host is { } host &&
            !string.Equals(host, "graph.microsoft.com", StringComparison.OrdinalIgnoreCase) ? host : null;
        _azureCli = FlowAgentRuntime.ResolveAzureCli();
    }

    public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
        GetTokenAsync(requestContext, cancellationToken).AsTask().GetAwaiter().GetResult();

    public override async ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        var resource = RequireResource(requestContext.Scopes);
        if (!string.Equals(resource.Host, "graph.microsoft.com", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(resource.Host, _sharePointHostname, StringComparison.OrdinalIgnoreCase))
            throw new BoundaryViolationException("Azure CLI credential was asked for an unapproved resource audience.");

        var tenantArgument = _tenantId is null ? string.Empty : $" --tenant {_tenantId}";
        var command = $"\"\"{_azureCli}\" account get-access-token{tenantArgument} --resource {resource.GetLeftPart(UriPartial.Authority)} --output json\"";
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe",
                Arguments = "/d /s /c " + command,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };
        process.StartInfo.Environment["AZURE_CORE_ENABLE_BROKER_ON_WINDOWS"] = "false";
        if (!process.Start()) throw new AuthenticationFailedException("Azure CLI could not be started.");
        var outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var errorTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        var output = await outputTask.ConfigureAwait(false);
        var error = await errorTask.ConfigureAwait(false);
        if (process.ExitCode != 0)
            throw new AuthenticationFailedException("Azure CLI could not acquire the approved Microsoft resource session: " +
                FlowAgentSafety.SafeText(string.IsNullOrWhiteSpace(error) ? output : error));

        try
        {
            var document = JsonNode.Parse(output)?.AsObject() ?? throw new FormatException();
            var token = document["accessToken"]?.GetValue<string>();
            if (string.IsNullOrWhiteSpace(token)) throw new FormatException();
            var expires = ParseExpiry(document);
            return new AccessToken(token, expires);
        }
        catch (Exception exception) when (exception is not AuthenticationFailedException)
        {
            throw new AuthenticationFailedException("Azure CLI returned an invalid token response.", exception);
        }
    }

    internal static Uri RequireResource(IEnumerable<string> scopes)
    {
        Uri? resource = null;
        foreach (var scope in scopes)
        {
            if (!Uri.TryCreate(scope, UriKind.Absolute, out var parsed) ||
                !string.Equals(parsed.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
                !parsed.IsDefaultPort || !string.IsNullOrEmpty(parsed.UserInfo))
                throw new BoundaryViolationException("Azure CLI credential received a non-canonical resource scope.");
            var current = new Uri(parsed.GetLeftPart(UriPartial.Authority));
            if (resource is not null && !string.Equals(resource.Host, current.Host, StringComparison.OrdinalIgnoreCase))
                throw new BoundaryViolationException("Azure CLI credential cannot combine resource audiences.");
            resource = current;
        }
        return resource ?? throw new BoundaryViolationException("Azure CLI credential requires one approved resource audience.");
    }

    private static DateTimeOffset ParseExpiry(JsonObject document)
    {
        if (document["expires_on"] is JsonValue unix && unix.TryGetValue<long>(out var seconds))
            return DateTimeOffset.FromUnixTimeSeconds(seconds);
        var text = document["expiresOn"]?.GetValue<string>();
        if (DateTimeOffset.TryParse(text, out var parsed)) return parsed;
        return DateTimeOffset.UtcNow.AddMinutes(30);
    }
}

internal sealed record DelegatedAuthenticationSession(
    TokenCredential Credential,
    string TenantId,
    string? Username,
    string Provider);
