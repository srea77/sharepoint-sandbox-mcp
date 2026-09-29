using System.Collections.Concurrent;
using System.IO.Compression;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using Azure.Core;
using Azure.Identity;
using SharePointMcp.Core;

namespace SharePointMcp.Infrastructure;

public sealed class OneDriveEvidenceService : IOneDriveEvidenceService
{
    private readonly TrustedSiteConfiguration _configuration;
    private readonly SiteBoundary _boundary;
    private readonly OneDriveScopeConfiguration _scope;
    private readonly TokenCredential _credential;
    private readonly HttpClient _http = new(new HttpClientHandler { AllowAutoRedirect = false }) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0/") };
    private static readonly string[] Scopes = ["https://graph.microsoft.com/Files.Read.All"];

    private OneDriveEvidenceService(TrustedSiteConfiguration configuration, SiteBoundary boundary, TokenCredential credential)
    {
        _configuration = configuration.Validate();
        _boundary = boundary;
        _scope = configuration.Scope.OneDrive ?? throw new ConfigurationException("OneDrive scope is not configured.");
        _credential = credential;
    }

    public static async Task<IOneDriveEvidenceService> TryStartAsync(TrustedSiteConfiguration configuration, SiteBoundary boundary, bool allowInteractive, CancellationToken cancellationToken = default)
    {
        if (configuration.Scope.OneDrive is null) return new UnavailableOneDriveEvidenceService("OneDrive drive/root scope was not configured during trusted setup.");
        try
        {
            var session = await AzureIdentityGraphTransport.CreateCredentialAsync(configuration.TenantId.ToString(), cancellationToken, Scopes, allowInteractive).ConfigureAwait(false);
            return new OneDriveEvidenceService(configuration, boundary, session.Credential);
        }
        catch (Exception exception) when (exception is AuthenticationFailedException or CredentialUnavailableException)
        {
            return new UnavailableOneDriveEvidenceService("A delegated OneDrive read session is unavailable: " + Safe(exception.Message));
        }
    }

    public Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(new JsonObject { ["available"] = true, ["driveId"] = _scope.DriveId, ["rootPath"] = _scope.RootPath, ["readOnly"] = true });
    }

    public async Task<JsonArray> SearchAsync(string query, int maxResults = 25, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.OneDriveRead);
        var safeQuery = RequireSearchQuery(query);
        if (maxResults is < 1 or > 100) throw new BoundaryViolationException("maxResults must be 1-100.");
        var root = RootPath();
        var search = Uri.EscapeDataString(safeQuery.Replace("'", "''", StringComparison.Ordinal));
        var path = root.Length == 0
            ? $"drives/{Uri.EscapeDataString(_scope.DriveId)}/root/search(q='{search}')?$top={maxResults}"
            : $"drives/{Uri.EscapeDataString(_scope.DriveId)}/root:/{root}:/search(q='{search}')?$top={maxResults}";
        var data = await SendJsonAsync(HttpMethod.Get, path, null, cancellationToken).ConfigureAwait(false);
        var result = new JsonArray();
        foreach (var item in (data?["value"] as JsonArray ?? []).OfType<JsonObject>())
        {
            RequireReturnedItemInsideRoot(item);
            result.Add(SanitizeItem(item));
        }
        return result;
    }

    public async Task<JsonObject> ReadAsync(string relativePath, int maxCharacters = 200_000, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.OneDriveRead);
        if (maxCharacters is < 1 or > 2_000_000) throw new BoundaryViolationException("maxCharacters must be 1-2,000,000.");
        var safePath = RequireRelativePath(relativePath);
        var combined = string.Join('/', new[] { RootPath(), safePath }.Where(value => value.Length > 0));
        var itemPath = $"drives/{Uri.EscapeDataString(_scope.DriveId)}/root:/{combined}";
        var metadata = await SendJsonAsync(HttpMethod.Get, itemPath, null, cancellationToken).ConfigureAwait(false) as JsonObject ?? new JsonObject();
        RequireReturnedItemInsideRoot(metadata);
        var bytes = await ReadContentAsync(itemPath + ":/content", cancellationToken).ConfigureAwait(false);
        if (bytes.LongLength > 100L * 1024 * 1024) throw new BoundaryViolationException("OneDrive evidence extraction is limited to 100 MB per file.");
        var text = ExtractText(bytes, Path.GetExtension(safePath).ToLowerInvariant());
        var truncated = text.Length > maxCharacters;
        if (truncated) text = text[..maxCharacters];
        return new JsonObject
        {
            ["relativePath"] = safePath,
            ["name"] = metadata["name"]?.DeepClone(),
            ["text"] = text,
            ["truncated"] = truncated,
            ["characterCountReturned"] = text.Length
        };
    }

    public ValueTask DisposeAsync()
    {
        _http.Dispose();
        return ValueTask.CompletedTask;
    }

    private async Task<JsonNode?> SendJsonAsync(HttpMethod method, string path, JsonNode? body, CancellationToken cancellationToken)
    {
        if (Uri.TryCreate(path, UriKind.Absolute, out _) || path.StartsWith('/') || path.Contains("..", StringComparison.Ordinal) || path.Contains('\\'))
            throw new BoundaryViolationException("OneDrive backend path escaped its fixed drive root.");
        var prefix = $"drives/{Uri.EscapeDataString(_scope.DriveId)}/";
        if (!path.StartsWith(prefix, StringComparison.Ordinal)) throw new BoundaryViolationException("OneDrive backend path escaped the configured drive.");
        var token = await _credential.GetTokenAsync(new TokenRequestContext(Scopes), cancellationToken).ConfigureAwait(false);
        using var request = new HttpRequestMessage(method, path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        if (body is not null) request.Content = new StringContent(body.ToJsonString(JsonDefaults.Options), Encoding.UTF8, "application/json");
        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        if ((int)response.StatusCode is >= 300 and < 400) throw new BoundaryViolationException("OneDrive JSON redirect was rejected.");
        var content = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode) throw new GraphBackendException($"Microsoft Graph returned {(int)response.StatusCode}: {Safe(content)}");
        return string.IsNullOrWhiteSpace(content) ? null : JsonNode.Parse(content);
    }

    private async Task<byte[]> ReadContentAsync(string path, CancellationToken cancellationToken)
    {
        var token = await _credential.GetTokenAsync(new TokenRequestContext(Scopes), cancellationToken).ConfigureAwait(false);
        using var request = new HttpRequestMessage(HttpMethod.Get, path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        if ((int)response.StatusCode is >= 300 and < 400)
        {
            var location = response.Headers.Location ?? throw new BoundaryViolationException("OneDrive content redirect omitted its location.");
            if (!location.IsAbsoluteUri || !string.Equals(location.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
                !(location.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase) || location.Host.EndsWith(".1drv.com", StringComparison.OrdinalIgnoreCase)))
                throw new BoundaryViolationException("OneDrive content redirect left Microsoft-controlled download origins.");
            using var downloadRequest = new HttpRequestMessage(HttpMethod.Get, location);
            using var downloadResponse = await _http.SendAsync(downloadRequest, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            if (!downloadResponse.IsSuccessStatusCode) throw new GraphBackendException($"OneDrive download returned {(int)downloadResponse.StatusCode}.");
            return await downloadResponse.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);
        }
        if (!response.IsSuccessStatusCode) throw new GraphBackendException($"OneDrive content returned {(int)response.StatusCode}.");
        return await response.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);
    }

    private void RequireReturnedItemInsideRoot(JsonObject item)
    {
        var path = item["parentReference"]?["path"]?.GetValue<string>() ?? string.Empty;
        var expected = $"/drives/{_scope.DriveId}/root:";
        if (!path.StartsWith(expected, StringComparison.OrdinalIgnoreCase)) throw new BoundaryViolationException("Microsoft Graph returned an item outside the configured OneDrive.");
        var root = RootPath();
        if (root.Length > 0 && !(path.Equals(expected + "/" + root, StringComparison.OrdinalIgnoreCase) || path.StartsWith(expected + "/" + root + "/", StringComparison.OrdinalIgnoreCase)))
            throw new BoundaryViolationException("Microsoft Graph returned an item outside the configured OneDrive root.");
    }

    private string RootPath() => _scope.RootPath.Trim('/');
    private static string RequireRelativePath(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 1024 || Uri.TryCreate(value, UriKind.Absolute, out _) || value.StartsWith('/') || value.StartsWith('\\'))
            throw new BoundaryViolationException("relativePath must be inside the configured OneDrive root.");
        var normalized = Uri.UnescapeDataString(value.Replace('\\', '/'));
        if (normalized.Split('/').Any(segment => segment is "" or "." or ".." || segment.Contains(':') || segment.Any(char.IsControl)))
            throw new BoundaryViolationException("relativePath contains an invalid or traversal segment.");
        return string.Join('/', normalized.Split('/').Select(Uri.EscapeDataString));
    }

    private static string RequireSearchQuery(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 256 || value.Any(char.IsControl) || value.Contains("://", StringComparison.Ordinal))
            throw new BoundaryViolationException("query must be non-empty plain text of at most 256 characters.");
        return value.Trim();
    }

    private static JsonObject SanitizeItem(JsonObject item) => new()
    {
        ["id"] = item["id"]?.DeepClone(),
        ["name"] = item["name"]?.DeepClone(),
        ["size"] = item["size"]?.DeepClone(),
        ["lastModifiedDateTime"] = item["lastModifiedDateTime"]?.DeepClone(),
        ["file"] = item["file"]?.DeepClone(),
        ["folder"] = item["folder"]?.DeepClone()
    };

    private static string ExtractText(byte[] bytes, string extension)
    {
        if (extension is ".txt" or ".csv" or ".json" or ".xml" or ".html" or ".htm" or ".md" or ".log") return Encoding.UTF8.GetString(bytes).Replace("\0", string.Empty, StringComparison.Ordinal);
        if (extension is not (".docx" or ".pptx" or ".xlsx")) throw new UnsupportedCapabilityException("OneDrive text extraction supports text and Open XML Word, PowerPoint, and Excel files; PDF extraction is not bundled.");
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entries = extension switch
        {
            ".docx" => archive.Entries.Where(entry => entry.FullName.StartsWith("word/", StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)),
            ".pptx" => archive.Entries.Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)),
            _ => archive.Entries.Where(entry => entry.FullName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase) || entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase))
        };
        var builder = new StringBuilder();
        foreach (var entry in entries.OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase))
        {
            using var input = entry.Open();
            var document = XDocument.Load(input, LoadOptions.None);
            foreach (var text in document.DescendantNodes().OfType<XText>().Select(node => node.Value).Where(text => !string.IsNullOrWhiteSpace(text))) builder.Append(text).Append(' ');
            builder.AppendLine();
        }
        return builder.ToString().Trim();
    }

    private static string Safe(string value)
    {
        var text = value.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return text.Length <= 600 ? text : text[..600];
    }
}

public sealed class UnavailableOneDriveEvidenceService(string reason) : IOneDriveEvidenceService
{
    public Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default) => Task.FromResult(new JsonObject { ["available"] = false, ["reason"] = reason });
    public Task<JsonArray> SearchAsync(string query, int maxResults = 25, CancellationToken cancellationToken = default) => Blocked<JsonArray>();
    public Task<JsonObject> ReadAsync(string relativePath, int maxCharacters = 200000, CancellationToken cancellationToken = default) => Blocked<JsonObject>();
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    private Task<T> Blocked<T>() => Task.FromException<T>(new UnsupportedCapabilityException(reason));
}

public sealed class GraphMailService : IMailService
{
    private readonly TrustedSiteConfiguration _configuration;
    private readonly SiteBoundary _boundary;
    private readonly MailScopeConfiguration _scope;
    private readonly IAuditLog _audit;
    private readonly TokenCredential _credential;
    private readonly string[] _scopes;
    private readonly HttpClient _http = new(new HttpClientHandler { AllowAutoRedirect = false }) { BaseAddress = new Uri("https://graph.microsoft.com/v1.0/") };
    private readonly ConcurrentDictionary<string, byte> _runtimeDrafts = new(StringComparer.Ordinal);

    private GraphMailService(TrustedSiteConfiguration configuration, SiteBoundary boundary, IAuditLog audit, TokenCredential credential, string[] scopes)
    {
        _configuration = configuration.Validate();
        _boundary = boundary;
        _scope = configuration.Scope.Mail ?? throw new ConfigurationException("Mail scope is not configured.");
        _audit = audit;
        _credential = credential;
        _scopes = scopes;
    }

    public static async Task<IMailService> TryStartAsync(TrustedSiteConfiguration configuration, SiteBoundary boundary, IAuditLog audit, bool allowInteractive, CancellationToken cancellationToken = default)
    {
        if (configuration.Scope.Mail is null) return new UnavailableMailService("Mail folder/recipient scope was not configured during trusted setup.");
        var scopes = new List<string>();
        if (configuration.Capabilities.MailRead) scopes.Add("https://graph.microsoft.com/Mail.Read");
        if (configuration.Capabilities.MailDraft) scopes.Add("https://graph.microsoft.com/Mail.ReadWrite");
        if (configuration.Capabilities.MailSend) scopes.Add("https://graph.microsoft.com/Mail.Send");
        if (scopes.Count == 0) return new UnavailableMailService("All trusted mail capabilities are disabled.");
        try
        {
            var session = await AzureIdentityGraphTransport.CreateCredentialAsync(configuration.TenantId.ToString(), cancellationToken, scopes.ToArray(), allowInteractive).ConfigureAwait(false);
            return new GraphMailService(configuration, boundary, audit, session.Credential, scopes.ToArray());
        }
        catch (Exception exception) when (exception is AuthenticationFailedException or CredentialUnavailableException)
        {
            return new UnavailableMailService("A delegated Outlook mail session is unavailable: " + Safe(exception.Message));
        }
    }

    public Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(new JsonObject
        {
            ["available"] = true,
            ["folderAliases"] = new JsonArray(_scope.FolderAliases.Keys.Select(value => (JsonNode?)JsonValue.Create(value)).ToArray()),
            ["allowedRecipientDomains"] = new JsonArray(_scope.AllowedRecipientDomains.Select(value => (JsonNode?)JsonValue.Create(value)).ToArray()),
            ["readEnabled"] = _configuration.Capabilities.MailRead,
            ["draftEnabled"] = _configuration.Capabilities.MailDraft,
            ["sendEnabled"] = _configuration.Capabilities.MailSend
        });
    }

    public async Task<JsonArray> SearchAsync(string folderAlias, string query, int maxResults = 25, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.MailRead);
        if (string.IsNullOrWhiteSpace(query) || query.Length > 256 || query.Any(char.IsControl)) throw new BoundaryViolationException("query must be 1-256 plain-text characters.");
        if (maxResults is < 1 or > 100) throw new BoundaryViolationException("maxResults must be 1-100.");
        var folder = ResolveFolder(folderAlias);
        var response = await SendAsync(HttpMethod.Get, $"me/mailFolders/{Uri.EscapeDataString(folder)}/messages?$select=id,subject,from,receivedDateTime,bodyPreview,hasAttachments,isRead&$top=100", null, cancellationToken).ConfigureAwait(false);
        return new JsonArray((response?["value"] as JsonArray ?? []).OfType<JsonObject>()
            .Where(message => new[] { message["subject"]?.ToString(), message["from"]?["emailAddress"]?["address"]?.ToString(), message["bodyPreview"]?.ToString() }
                .Any(value => value?.Contains(query, StringComparison.OrdinalIgnoreCase) is true))
            .Take(maxResults).Select(message => (JsonNode)SanitizeMessage(message)).ToArray());
    }

    public async Task<JsonObject> GetMessageAsync(string folderAlias, string messageId, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.MailRead);
        var folder = ResolveFolder(folderAlias);
        var id = RequireOpaqueGraphId(messageId, "messageId");
        var response = await SendAsync(HttpMethod.Get, $"me/mailFolders/{Uri.EscapeDataString(folder)}/messages/{Uri.EscapeDataString(id)}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,hasAttachments,internetMessageId", null, cancellationToken).ConfigureAwait(false);
        return SanitizeMessage(response as JsonObject ?? new JsonObject());
    }

    public async Task<JsonArray> ListAttachmentsAsync(string folderAlias, string messageId, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.MailRead);
        var folder = ResolveFolder(folderAlias);
        var id = RequireOpaqueGraphId(messageId, "messageId");
        _ = await GetMessageAsync(folderAlias, id, cancellationToken).ConfigureAwait(false);
        var response = await SendAsync(HttpMethod.Get, $"me/mailFolders/{Uri.EscapeDataString(folder)}/messages/{Uri.EscapeDataString(id)}/attachments?$select=id,name,contentType,size,isInline,lastModifiedDateTime", null, cancellationToken).ConfigureAwait(false);
        return response?["value"] as JsonArray ?? [];
    }

    public async Task<JsonObject> ReadAttachmentAsync(string folderAlias, string messageId, string attachmentId, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.MailRead);
        var folder = ResolveFolder(folderAlias);
        var message = RequireOpaqueGraphId(messageId, "messageId");
        var attachment = RequireOpaqueGraphId(attachmentId, "attachmentId");
        _ = await GetMessageAsync(folderAlias, message, cancellationToken).ConfigureAwait(false);
        var response = await SendAsync(HttpMethod.Get, $"me/mailFolders/{Uri.EscapeDataString(folder)}/messages/{Uri.EscapeDataString(message)}/attachments/{Uri.EscapeDataString(attachment)}", null, cancellationToken).ConfigureAwait(false) as JsonObject ?? new JsonObject();
        if (response["contentBytes"]?.GetValue<string>() is { } content && content.Length > 40_000_000) throw new BoundaryViolationException("Mail attachment exceeds the 30 MB MCP transfer limit.");
        return response;
    }

    public async Task<MutationResult> CreateDraftAsync(MailDraftDefinition definition, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.MailDraft);
        ValidateDraft(definition);
        if (dryRun) return new("mail_draft_create", definition.Subject, false, "Dry run: validated recipients and draft content against trusted mail policy.");
        var body = new JsonObject
        {
            ["subject"] = definition.Subject,
            ["body"] = new JsonObject { ["contentType"] = definition.BodyIsHtml ? "HTML" : "Text", ["content"] = definition.Body },
            ["toRecipients"] = Recipients(definition.To),
            ["ccRecipients"] = Recipients(definition.Cc ?? [])
        };
        var response = await SendAsync(HttpMethod.Post, "me/messages", body, cancellationToken).ConfigureAwait(false) as JsonObject ?? new JsonObject();
        var id = response["id"]?.GetValue<string>() ?? throw new GraphBackendException("Microsoft Graph created a draft without returning its ID.");
        _runtimeDrafts.TryAdd(id, 0);
        return await Audited(new("mail_draft_create", definition.Subject, true, "Created an unsent draft; its ID is usable only by this runtime's send tool.", Data: SanitizeMessage(response)), cancellationToken).ConfigureAwait(false);
    }

    public async Task<MutationResult> SendDraftAsync(string draftId, bool dryRun = true, CancellationToken cancellationToken = default)
    {
        _boundary.RequireCapability(RuntimeCapability.MailSend);
        var id = RequireOpaqueGraphId(draftId, "draftId");
        if (!_runtimeDrafts.ContainsKey(id)) throw new BoundaryViolationException("draftId was not created by this trusted runtime.");
        if (dryRun) return new("mail_draft_send", id, false, "Dry run: would send this runtime-created draft.");
        _ = await SendAsync(HttpMethod.Post, $"me/messages/{Uri.EscapeDataString(id)}/send", new JsonObject(), cancellationToken).ConfigureAwait(false);
        _runtimeDrafts.TryRemove(id, out _);
        return await Audited(new("mail_draft_send", id, true, "Sent the runtime-created draft."), cancellationToken).ConfigureAwait(false);
    }

    public ValueTask DisposeAsync()
    {
        _http.Dispose();
        return ValueTask.CompletedTask;
    }

    private string ResolveFolder(string aliasOrId)
    {
        var safe = RequireOpaqueGraphId(aliasOrId, "folderAlias");
        foreach (var pair in _scope.FolderAliases)
            if (string.Equals(pair.Key, safe, StringComparison.OrdinalIgnoreCase)) return pair.Value;
        return _scope.AllowedFolderIds.FirstOrDefault(value => string.Equals(value, safe, StringComparison.Ordinal))
            ?? throw new BoundaryViolationException("Mail folder is outside trusted configured scope.");
    }

    private void ValidateDraft(MailDraftDefinition definition)
    {
        if (string.IsNullOrWhiteSpace(definition.Subject) || definition.Subject.Length > 998 || definition.Subject.Any(character => character is '\r' or '\n' or '\0')) throw new BoundaryViolationException("Mail subject is invalid.");
        if (definition.Body.Length > 2_000_000 || definition.Body.Contains("javascript:", StringComparison.OrdinalIgnoreCase)) throw new BoundaryViolationException("Mail body is invalid or too large.");
        var recipients = definition.To.Concat(definition.Cc ?? []).ToArray();
        if (recipients.Length is < 1 or > 100) throw new BoundaryViolationException("Mail draft must have 1-100 recipients.");
        foreach (var recipient in recipients)
        {
            var at = recipient.LastIndexOf('@');
            if (at <= 0 || at == recipient.Length - 1 || recipient.Any(char.IsControl)) throw new BoundaryViolationException("A mail recipient is invalid.");
            var domain = recipient[(at + 1)..];
            if (!_scope.AllowedRecipientDomains.Any(value => string.Equals(value, domain, StringComparison.OrdinalIgnoreCase)))
                throw new BoundaryViolationException($"Recipient domain '{domain}' is outside trusted configured scope.");
        }
    }

    private static JsonArray Recipients(IEnumerable<string> values) => new(values.Select(value => (JsonNode)new JsonObject { ["emailAddress"] = new JsonObject { ["address"] = value } }).ToArray());

    private async Task<JsonNode?> SendAsync(HttpMethod method, string path, JsonNode? body, CancellationToken cancellationToken)
    {
        if (Uri.TryCreate(path, UriKind.Absolute, out _) || path.StartsWith('/') || path.Contains("..", StringComparison.Ordinal) || path.Contains('\\') || !path.StartsWith("me/", StringComparison.Ordinal))
            throw new BoundaryViolationException("Mail backend path escaped the signed-in user's fixed mailbox resource.");
        var token = await _credential.GetTokenAsync(new TokenRequestContext(_scopes), cancellationToken).ConfigureAwait(false);
        using var request = new HttpRequestMessage(method, path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        if (body is not null) request.Content = new StringContent(body.ToJsonString(JsonDefaults.Options), Encoding.UTF8, "application/json");
        using var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        if ((int)response.StatusCode is >= 300 and < 400) throw new BoundaryViolationException("Mail redirect was rejected.");
        var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode) throw new GraphBackendException($"Microsoft Graph mail returned {(int)response.StatusCode}: {Safe(text)}");
        return string.IsNullOrWhiteSpace(text) ? null : JsonNode.Parse(text);
    }

    private static string RequireOpaqueGraphId(string value, string name)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 1024 || value.Any(char.IsControl) || value.Contains("://", StringComparison.Ordinal))
            throw new BoundaryViolationException($"{name} is not a safe opaque identifier.");
        return value;
    }

    private static JsonObject SanitizeMessage(JsonObject message)
    {
        var clone = message.DeepClone().AsObject();
        foreach (var key in clone.Select(pair => pair.Key).Where(key => key.Contains("token", StringComparison.OrdinalIgnoreCase)).ToArray()) clone.Remove(key);
        return clone;
    }

    private async Task<MutationResult> Audited(MutationResult result, CancellationToken cancellationToken)
    {
        await _audit.WriteAsync(new(DateTimeOffset.UtcNow, result.Operation, _configuration.CanonicalSiteIdentifier, result.Target, result.Outcome, result.CorrelationId), cancellationToken).ConfigureAwait(false);
        return result;
    }

    private static string Safe(string value)
    {
        var text = value.Replace('\r', ' ').Replace('\n', ' ').Trim();
        foreach (var marker in new[] { "Bearer ", "access_token", "refresh_token", "cookie" }) if (text.Contains(marker, StringComparison.OrdinalIgnoreCase)) return "[sensitive backend detail redacted]";
        return text.Length <= 600 ? text : text[..600];
    }
}

public sealed class UnavailableMailService(string reason) : IMailService
{
    public Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default) => Task.FromResult(new JsonObject { ["available"] = false, ["reason"] = reason });
    public Task<JsonArray> SearchAsync(string folderAlias, string query, int maxResults = 25, CancellationToken cancellationToken = default) => Blocked<JsonArray>();
    public Task<JsonObject> GetMessageAsync(string folderAlias, string messageId, CancellationToken cancellationToken = default) => Blocked<JsonObject>();
    public Task<JsonArray> ListAttachmentsAsync(string folderAlias, string messageId, CancellationToken cancellationToken = default) => Blocked<JsonArray>();
    public Task<JsonObject> ReadAttachmentAsync(string folderAlias, string messageId, string attachmentId, CancellationToken cancellationToken = default) => Blocked<JsonObject>();
    public Task<MutationResult> CreateDraftAsync(MailDraftDefinition definition, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public Task<MutationResult> SendDraftAsync(string draftId, bool dryRun = true, CancellationToken cancellationToken = default) => Blocked<MutationResult>();
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    private Task<T> Blocked<T>() => Task.FromException<T>(new UnsupportedCapabilityException(reason));
}

public static class ExternalScopeBootstrapper
{
    public static async Task<JsonObject> DiscoverAsync(
        Guid tenantId,
        string service,
        CancellationToken cancellationToken = default)
    {
        var normalized = service.ToLowerInvariant() switch
        {
            "onedrive" => "onedrive",
            "mail" => "mail",
            _ => throw new ArgumentException("--service must be onedrive or mail.")
        };
        var scopes = normalized == "onedrive"
            ? new[] { "https://graph.microsoft.com/Files.Read.All" }
            : new[] { "https://graph.microsoft.com/Mail.Read" };
        var path = normalized == "onedrive"
            ? "me/drives?$select=id,name,driveType,webUrl"
            : "me/mailFolders?$select=id,displayName,parentFolderId,childFolderCount&$top=100";
        var session = await AzureIdentityGraphTransport.CreateCredentialAsync(
            tenantId.ToString(), cancellationToken, scopes, allowInteractive: true).ConfigureAwait(false);
        var token = await session.Credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken).ConfigureAwait(false);
        using var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false })
        {
            BaseAddress = new Uri("https://graph.microsoft.com/v1.0/")
        };
        using var request = new HttpRequestMessage(HttpMethod.Get, path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        using var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        if ((int)response.StatusCode is >= 300 and < 400) throw new BoundaryViolationException("External-scope bootstrap discovery redirect was rejected.");
        var text = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode) throw new GraphBackendException($"Microsoft Graph bootstrap discovery returned {(int)response.StatusCode}: {SafeDiscovery(text)}");
        var values = JsonNode.Parse(text)?["value"] as JsonArray ?? [];
        var resources = new JsonArray(values.OfType<JsonObject>().Select(value => normalized == "onedrive"
            ? (JsonNode)new JsonObject
            {
                ["id"] = value["id"]?.DeepClone(),
                ["name"] = value["name"]?.DeepClone(),
                ["driveType"] = value["driveType"]?.DeepClone(),
                ["webUrl"] = value["webUrl"]?.DeepClone()
            }
            : new JsonObject
            {
                ["id"] = value["id"]?.DeepClone(),
                ["displayName"] = value["displayName"]?.DeepClone(),
                ["parentFolderId"] = value["parentFolderId"]?.DeepClone(),
                ["childFolderCount"] = value["childFolderCount"]?.DeepClone()
            }).ToArray());
        return new JsonObject
        {
            ["resourceType"] = normalized == "onedrive" ? "drive" : "mailFolder",
            ["resources"] = resources,
            ["note"] = "Trusted bootstrap discovery only. Copy selected IDs into a reviewed runtime policy; this operation is not exposed through MCP."
        };
    }

    private static string SafeDiscovery(string value)
    {
        if (new[] { "Bearer ", "access_token", "refresh_token", "cookie" }.Any(marker => value.Contains(marker, StringComparison.OrdinalIgnoreCase)))
            return "[sensitive backend detail redacted]";
        var normalized = value.Replace('\r', ' ').Replace('\n', ' ').Trim();
        return normalized.Length <= 800 ? normalized : normalized[..800];
    }
}
