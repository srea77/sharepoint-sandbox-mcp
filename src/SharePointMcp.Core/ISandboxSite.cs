using System.Text.Json.Nodes;

namespace SharePointMcp.Core;

public interface ISandboxSite
{
    TrustedSiteConfiguration Configuration { get; }

    Task<SiteInfo> GetInfoAsync(CancellationToken cancellationToken = default);
    Task<SiteIdentity> GetIdentityAsync(CancellationToken cancellationToken = default);
    Task<JsonObject> GetSiteAppearanceAsync(CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateSiteAppearanceAsync(SiteAppearanceUpdate definition, string? expectedETag = null, bool dryRun = true, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<CapabilityStatus>> GetCapabilitiesAsync(CancellationToken cancellationToken = default);
    Task<SiteInventory> GetInventoryAsync(CancellationToken cancellationToken = default);
    Task<HealthReport> ValidateHealthAsync(CancellationToken cancellationToken = default);

    Task<JsonObject> GetThemeAsync(CancellationToken cancellationToken = default);
    Task<JsonArray> ListAvailableThemesAsync(CancellationToken cancellationToken = default);
    Task<SiteThemeValidationResult> ValidateThemeAsync(SiteThemeDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> ApplyThemeAsync(SiteThemeDefinition definition, bool dryRun = true, CancellationToken cancellationToken = default);

    Task<JsonArray> ListListsAsync(ListKind? kind = null, CancellationToken cancellationToken = default);
    Task<JsonObject> GetListAsync(string listName, CancellationToken cancellationToken = default);
    Task<MutationResult> CreateListAsync(ListDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateListAsync(string listName, ListDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> DeleteListAsync(string listName, bool dryRun = false, CancellationToken cancellationToken = default);

    Task<JsonArray> ListColumnsAsync(string listName, CancellationToken cancellationToken = default);
    Task<JsonObject> GetColumnAsync(string listName, string columnName, CancellationToken cancellationToken = default);
    Task<MutationResult> CreateColumnAsync(string listName, ColumnDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateColumnAsync(string listName, string columnName, ColumnDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> DeleteColumnAsync(string listName, string columnName, bool dryRun = false, CancellationToken cancellationToken = default);

    Task<JsonObject> GetItemAsync(string listName, string itemId, CancellationToken cancellationToken = default);
    Task<JsonArray> QueryItemsAsync(string listName, ItemQuery query, CancellationToken cancellationToken = default);
    Task<ItemQueryResult> QueryItemsPageAsync(string listName, ItemQuery query, CancellationToken cancellationToken = default);
    Task<MutationResult> CreateItemAsync(string listName, JsonObject fields, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateItemAsync(string listName, string itemId, JsonObject fields, string? expectedETag = null, CancellationToken cancellationToken = default);
    Task<MutationResult> DeleteItemAsync(string listName, string itemId, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> RecycleItemAsync(string listName, string itemId, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<JsonArray> GetItemVersionsAsync(string listName, string itemId, CancellationToken cancellationToken = default);
    Task<MutationResult> RestoreItemVersionAsync(string listName, string itemId, string versionId, bool dryRun = false, CancellationToken cancellationToken = default);

    Task<JsonArray> ListViewsAsync(string listName, CancellationToken cancellationToken = default);
    Task<JsonObject> GetViewAsync(string listName, string viewName, CancellationToken cancellationToken = default);
    Task<MutationResult> CreateViewAsync(string listName, ViewDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateViewAsync(string listName, string viewName, ViewDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> DeleteViewAsync(string listName, string viewName, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> SetDefaultViewAsync(string listName, string viewName, CancellationToken cancellationToken = default);

    Task<JsonObject> GetFormAsync(string listName, CancellationToken cancellationToken = default);
    Task<MutationResult> ConfigureFormAsync(string listName, FormDefinition definition, CancellationToken cancellationToken = default);
    Task<JsonArray> GetFieldVisibilityRulesAsync(string listName, CancellationToken cancellationToken = default);
    Task<MutationResult> SetFieldVisibilityRuleAsync(string listName, FieldVisibilityRule rule, CancellationToken cancellationToken = default);
    Task<MutationResult> SetColumnFormattingAsync(string listName, string columnName, JsonNode formatting, CancellationToken cancellationToken = default);
    Task<MutationResult> ClearColumnFormattingAsync(string listName, string columnName, CancellationToken cancellationToken = default);
    Task<MutationResult> SetViewFormattingAsync(string listName, string viewName, JsonNode formatting, CancellationToken cancellationToken = default);
    Task<MutationResult> ClearViewFormattingAsync(string listName, string viewName, CancellationToken cancellationToken = default);

    Task<JsonArray> ListFilesAsync(string relativePath, CancellationToken cancellationToken = default);
    Task<JsonObject> GetFileMetadataAsync(string relativePath, CancellationToken cancellationToken = default);
    Task<byte[]> ReadFileAsync(string relativePath, CancellationToken cancellationToken = default);
    Task<MutationResult> WriteFileAsync(string relativePath, byte[] content, bool overwrite, CancellationToken cancellationToken = default);
    Task<MutationResult> DeleteFileAsync(string relativePath, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> MoveFileAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default);
    Task<MutationResult> CopyFileAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default);
    Task<MutationResult> CopyFileIfAbsentAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default);
    Task<MutationResult> RecycleFileAsync(string relativePath, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<JsonArray> GetFileVersionsAsync(string relativePath, CancellationToken cancellationToken = default);
    Task<MutationResult> RestoreFileVersionAsync(string relativePath, string versionId, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> CreateFolderAsync(string relativePath, CancellationToken cancellationToken = default);
    Task<MutationResult> DeleteFolderAsync(string relativePath, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> MoveFolderAsync(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken = default);
    Task<JsonArray> SearchEvidenceAsync(string query, int maxResults = 25, bool includeContent = false, CancellationToken cancellationToken = default);
    Task<JsonObject> ReadExtractedTextAsync(string relativePath, int maxCharacters = 200_000, CancellationToken cancellationToken = default);

    Task<JsonArray> ListContentTypesAsync(string? listName = null, CancellationToken cancellationToken = default);
    Task<MutationResult> CreateContentTypeAsync(ContentTypeDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateContentTypeAsync(string contentTypeId, ContentTypeDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> AddContentTypeToListAsync(string listName, string contentTypeId, CancellationToken cancellationToken = default);

    Task<JsonArray> ListPagesAsync(CancellationToken cancellationToken = default);
    Task<JsonObject> GetPageAsync(string pageName, CancellationToken cancellationToken = default);
    Task<MutationResult> CreatePageAsync(PageDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdatePageAsync(string pageName, PageDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> DeletePageAsync(string pageName, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> PublishPageAsync(string pageName, CancellationToken cancellationToken = default);
    Task<MutationResult> UnpublishPageAsync(string pageName, CancellationToken cancellationToken = default);
    Task<MutationResult> SetHomePageAsync(string pageName, CancellationToken cancellationToken = default);
    Task<MutationResult> CopyPageIfAbsentAsync(string sourcePageName, string destinationPageName, CancellationToken cancellationToken = default);
    Task<MutationResult> SavePageDraftAsync(string pageName, PageDefinition definition, string? expectedETag = null, CancellationToken cancellationToken = default);
    Task<JsonArray> GetPageVersionsAsync(string pageName, CancellationToken cancellationToken = default);
    Task<MutationResult> RestorePageVersionAsync(string pageName, string versionLabel, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> RecyclePageAsync(string pageName, bool dryRun = false, CancellationToken cancellationToken = default);

    Task<JsonArray> GetAvailableWebPartsAsync(CancellationToken cancellationToken = default);
    Task<JsonObject> GetWebPartAsync(string pageName, string webPartId, CancellationToken cancellationToken = default);
    Task<MutationResult> AddWebPartAsync(string pageName, string sectionKey, WebPartDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateWebPartAsync(string pageName, string webPartId, WebPartDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> MoveWebPartAsync(string pageName, string webPartId, string sectionKey, int column, int order, CancellationToken cancellationToken = default);
    Task<MutationResult> DeleteWebPartAsync(string pageName, string webPartId, bool dryRun = false, CancellationToken cancellationToken = default);

    Task<JsonArray> GetNavigationAsync(NavigationLocation location = NavigationLocation.QuickLaunch, CancellationToken cancellationToken = default);
    Task<MutationResult> AddNavigationAsync(NavigationDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> UpdateNavigationAsync(string key, NavigationDefinition definition, CancellationToken cancellationToken = default);
    Task<MutationResult> RemoveNavigationAsync(string key, NavigationLocation location = NavigationLocation.QuickLaunch, bool dryRun = false, CancellationToken cancellationToken = default);
    Task<MutationResult> ReorderNavigationAsync(IReadOnlyList<string> keys, NavigationLocation location = NavigationLocation.QuickLaunch, string? parentKey = null, CancellationToken cancellationToken = default);

    Task<JsonObject> InspectPermissionsAsync(PermissionTarget? target = null, CancellationToken cancellationToken = default);
    Task<MutationResult> GrantPermissionAsync(string principal, string role, PermissionTarget? target = null, CancellationToken cancellationToken = default);
    Task<MutationResult> RevokePermissionAsync(string principal, string role, PermissionTarget? target = null, CancellationToken cancellationToken = default);
    Task<MutationResult> ResetPermissionInheritanceAsync(PermissionTarget target, CancellationToken cancellationToken = default);
}
