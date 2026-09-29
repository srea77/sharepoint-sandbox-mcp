using System.Text.Json;
using ModelContextProtocol.Server;
using SharePointMcp.Core;

namespace SharePointMcp.Server;

[McpServerToolType]
public sealed class ExternalEvidenceTools(IOneDriveEvidenceService oneDrive, IMailService mail)
{
    [McpServerTool(Name = "onedrive_get_status", ReadOnly = true), System.ComponentModel.Description("Report the immutable OneDrive drive/root scope and authentication availability.")]
    public async Task<string> OneDriveStatus(CancellationToken cancellationToken) => Json(await oneDrive.GetStatusAsync(cancellationToken));

    [McpServerTool(Name = "onedrive_search", ReadOnly = true), System.ComponentModel.Description("Search only beneath the immutable configured OneDrive drive/root.")]
    public async Task<string> OneDriveSearch(string query, int maxResults = 25, CancellationToken cancellationToken = default) => Json(await oneDrive.SearchAsync(query, maxResults, cancellationToken));

    [McpServerTool(Name = "onedrive_read_text", ReadOnly = true), System.ComponentModel.Description("Read and locally extract text from an exact path beneath the immutable configured OneDrive root.")]
    public async Task<string> OneDriveReadText(string relativePath, int maxCharacters = 200000, CancellationToken cancellationToken = default) => Json(await oneDrive.ReadAsync(relativePath, maxCharacters, cancellationToken));

    [McpServerTool(Name = "mail_get_status", ReadOnly = true), System.ComponentModel.Description("Report configured mail folder aliases, recipient-domain policy, and separately enabled read/draft/send capabilities.")]
    public async Task<string> MailStatus(CancellationToken cancellationToken) => Json(await mail.GetStatusAsync(cancellationToken));

    [McpServerTool(Name = "mail_search", ReadOnly = true), System.ComponentModel.Description("Search message metadata only in a trusted configured mail-folder alias.")]
    public async Task<string> MailSearch(string folderAlias, string query, int maxResults = 25, CancellationToken cancellationToken = default) => Json(await mail.SearchAsync(folderAlias, query, maxResults, cancellationToken));

    [McpServerTool(Name = "mail_get", ReadOnly = true), System.ComponentModel.Description("Read one message that is proven to belong to a trusted configured mail folder.")]
    public async Task<string> MailGet(string folderAlias, string messageId, CancellationToken cancellationToken) => Json(await mail.GetMessageAsync(folderAlias, messageId, cancellationToken));

    [McpServerTool(Name = "mail_attachment_list", ReadOnly = true), System.ComponentModel.Description("List attachments on one message in a trusted configured mail folder.")]
    public async Task<string> MailAttachmentList(string folderAlias, string messageId, CancellationToken cancellationToken) => Json(await mail.ListAttachmentsAsync(folderAlias, messageId, cancellationToken));

    [McpServerTool(Name = "mail_attachment_read", ReadOnly = true), System.ComponentModel.Description("Read one attachment from a message in a trusted configured mail folder.")]
    public async Task<string> MailAttachmentRead(string folderAlias, string messageId, string attachmentId, CancellationToken cancellationToken) => Json(await mail.ReadAttachmentAsync(folderAlias, messageId, attachmentId, cancellationToken));

    [McpServerTool(Name = "mail_draft_create", Idempotent = true), System.ComponentModel.Description("Create an unsent draft only when mailDraft is enabled and every recipient domain is trusted; dryRun defaults true.")]
    public async Task<string> MailDraftCreate(MailDraftDefinition definition, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await mail.CreateDraftAsync(definition, dryRun, cancellationToken));

    [McpServerTool(Name = "mail_draft_send", Idempotent = true), System.ComponentModel.Description("Send only a draft created by this trusted runtime when mailSend is separately enabled; dryRun defaults true.")]
    public async Task<string> MailDraftSend(string draftId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await mail.SendDraftAsync(draftId, dryRun, cancellationToken));

    private static string Json<T>(T value) => JsonSerializer.Serialize(value, JsonDefaults.Options);
}
