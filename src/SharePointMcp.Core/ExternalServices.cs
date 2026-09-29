using System.Text.Json.Nodes;

namespace SharePointMcp.Core;

public sealed record MailDraftDefinition(
    string Subject,
    string Body,
    bool BodyIsHtml,
    IReadOnlyList<string> To,
    IReadOnlyList<string>? Cc = null);

public interface IOneDriveEvidenceService : IAsyncDisposable
{
    Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default);
    Task<JsonArray> SearchAsync(string query, int maxResults = 25, CancellationToken cancellationToken = default);
    Task<JsonObject> ReadAsync(string relativePath, int maxCharacters = 200_000, CancellationToken cancellationToken = default);
}
public interface IMailService : IAsyncDisposable
{
    Task<JsonObject> GetStatusAsync(CancellationToken cancellationToken = default);
    Task<JsonArray> SearchAsync(string folderAlias, string query, int maxResults = 25, CancellationToken cancellationToken = default);
    Task<JsonObject> GetMessageAsync(string folderAlias, string messageId, CancellationToken cancellationToken = default);
    Task<JsonArray> ListAttachmentsAsync(string folderAlias, string messageId, CancellationToken cancellationToken = default);
    Task<JsonObject> ReadAttachmentAsync(string folderAlias, string messageId, string attachmentId, CancellationToken cancellationToken = default);
    Task<MutationResult> CreateDraftAsync(MailDraftDefinition definition, bool dryRun = true, CancellationToken cancellationToken = default);
    Task<MutationResult> SendDraftAsync(string draftId, bool dryRun = true, CancellationToken cancellationToken = default);
}
