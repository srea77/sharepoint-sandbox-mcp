using System.ComponentModel;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Server;
using SharePointMcp.Core;

namespace SharePointMcp.Server;

[McpServerToolType]
public sealed class SharePointTools(ISandboxSite site, BlueprintEngine blueprints, BlueprintValidator blueprintValidator)
{
    [McpServerTool(Name = "who_am_i", ReadOnly = true), Description("Show the delegated Microsoft identity as seen by the configured SharePoint site, including site-effective permission bits but no authentication tokens.")]
    public async Task<string> WhoAmI(CancellationToken cancellationToken) => Json(await site.GetIdentityAsync(cancellationToken));

    [McpServerTool(Name = "get_configured_scope", ReadOnly = true), Description("Show the immutable trusted runtime scope and capability switches without exposing credentials or allowing configuration changes.")]
    public string GetConfiguredScope()
    {
        var configuration = site.Configuration;
        return Json(new
        {
            site = new
            {
                configuration.TenantHostname,
                siteUrl = configuration.SiteUrl.AbsoluteUri,
                configuration.GraphSiteId,
                configuration.SiteCollectionId,
                configuration.WebId
            },
            configuration.Capabilities,
            resources = new
            {
                configuration.Scope.AllowedLists,
                configuration.Scope.AllowedLibraries,
                configuration.Scope.AllowedPages,
                configuration.Scope.ListAliases,
                configuration.Scope.LibraryAliases,
                configuration.Scope.PageAliases,
                powerAutomate = configuration.Scope.PowerAutomate is null ? null : new
                {
                    configuration.Scope.PowerAutomate.EnvironmentId,
                    configuration.Scope.PowerAutomate.FlowAliases,
                    configuration.Scope.PowerAutomate.AllowedNewFlowNames,
                    configuration.Scope.PowerAutomate.AllowedConnectorNames
                },
                oneDrive = configuration.Scope.OneDrive is null ? null : new
                {
                    configuration.Scope.OneDrive.DriveId,
                    configuration.Scope.OneDrive.RootPath
                },
                mail = configuration.Scope.Mail is null ? null : new
                {
                    configuration.Scope.Mail.AllowedFolderIds,
                    configuration.Scope.Mail.AllowedRecipientDomains
                }
            }
        });
    }

    [McpServerTool(Name = "site_get_info", ReadOnly = true), Description("Get the immutable configured SharePoint sandbox site's identity and metadata. There is no site selector.")]
    public async Task<string> SiteGetInfo(CancellationToken cancellationToken) => Json(await site.GetInfoAsync(cancellationToken));

    [McpServerTool(Name = "site_get_capabilities", ReadOnly = true), Description("Report implemented, partial, experimental, and blocked capability families for the configured site.")]
    public async Task<string> SiteGetCapabilities(CancellationToken cancellationToken) => Json(await site.GetCapabilitiesAsync(cancellationToken));

    [McpServerTool(Name = "site_inventory", ReadOnly = true), Description("Inventory lists, libraries, pages, and other supported artifacts only in the configured site.")]
    public async Task<string> SiteInventory(CancellationToken cancellationToken) => Json(await site.GetInventoryAsync(cancellationToken));

    [McpServerTool(Name = "site_validate_health", ReadOnly = true), Description("Validate trusted configuration, authentication, and connectivity to the configured site.")]
    public async Task<string> SiteValidateHealth(CancellationToken cancellationToken) => Json(await site.ValidateHealthAsync(cancellationToken));

    [McpServerTool(Name = "site_appearance_get", ReadOnly = true), Description("Read title, description, logo, header, footer, navigation style, audience targeting, welcome page, and ETag from the configured site.")]
    public async Task<string> SiteAppearanceGet(CancellationToken cancellationToken) => Json(await site.GetSiteAppearanceAsync(cancellationToken));

    [McpServerTool(Name = "site_appearance_update", Idempotent = true), Description("Update typed appearance settings on the configured site. Requires trusted siteSettingsWrite capability; dryRun defaults true and expectedETag enables optimistic concurrency.")]
    public async Task<string> SiteAppearanceUpdate(SiteAppearanceUpdate definition, string? expectedETag = null, bool dryRun = true, CancellationToken cancellationToken = default) =>
        Json(await site.UpdateSiteAppearanceAsync(definition, expectedETag, dryRun, cancellationToken));

    [McpServerTool(Name = "site_theme_get", ReadOnly = true), Description("Read live theme information and the trusted last-applied site-theme state for the configured site only.")]
    public async Task<string> SiteThemeGet(CancellationToken cancellationToken) => Json(await site.GetThemeAsync(cancellationToken));

    [McpServerTool(Name = "site_theme_list", ReadOnly = true), Description("List themes available to the configured site without exposing tenant-wide theme mutation.")]
    public async Task<string> SiteThemeList(CancellationToken cancellationToken) => Json(await site.ListAvailableThemesAsync(cancellationToken));

    [McpServerTool(Name = "site_theme_validate", ReadOnly = true), Description("Validate and normalize a typed semantic SharePoint palette without mutation.")]
    public async Task<string> SiteThemeValidate(SiteThemeDefinition definition, CancellationToken cancellationToken) => Json(await site.ValidateThemeAsync(definition, cancellationToken));

    [McpServerTool(Name = "site_theme_apply", Idempotent = true), Description("Apply a validated palette to the configured site only through the exact site-scoped ApplyTheme endpoint; dryRun defaults true.")]
    public async Task<string> SiteThemeApply(SiteThemeDefinition definition, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.ApplyThemeAsync(definition, dryRun, cancellationToken));

    [McpServerTool(Name = "list_list", ReadOnly = true), Description("List SharePoint lists in the configured site (document libraries excluded).")]
    public async Task<string> ListList(CancellationToken cancellationToken) => Json(await site.ListListsAsync(ListKind.GenericList, cancellationToken));

    [McpServerTool(Name = "list_get", ReadOnly = true), Description("Get a list by site-local title.")]
    public async Task<string> ListGet([Description("Site-local list title; never a URL or site ID.")] string listName, CancellationToken cancellationToken) => Json(await site.GetListAsync(listName, cancellationToken));

    [McpServerTool(Name = "list_create", Idempotent = true), Description("Create a generic SharePoint list in the configured site.")]
    public async Task<string> ListCreate(ListDefinition definition, CancellationToken cancellationToken) => Json(await site.CreateListAsync(definition with { Kind = ListKind.GenericList }, cancellationToken));

    [McpServerTool(Name = "list_update", Idempotent = true), Description("Update Graph-supported settings of a list in the configured site.")]
    public async Task<string> ListUpdate(string listName, ListDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdateListAsync(listName, definition with { Kind = ListKind.GenericList }, cancellationToken));

    [McpServerTool(Name = "list_delete", Destructive = true, Idempotent = true), Description("Delete a list and all its data; dryRun defaults true.")]
    public async Task<string> ListDelete(string listName, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteListAsync(listName, dryRun, cancellationToken));

    [McpServerTool(Name = "column_list", ReadOnly = true), Description("List columns in one configured-site list.")]
    public async Task<string> ColumnList(string listName, CancellationToken cancellationToken) => Json(await site.ListColumnsAsync(listName, cancellationToken));

    [McpServerTool(Name = "column_get", ReadOnly = true), Description("Get one column by site-local internal or display name.")]
    public async Task<string> ColumnGet(string listName, string columnName, CancellationToken cancellationToken) => Json(await site.GetColumnAsync(listName, columnName, cancellationToken));

    [McpServerTool(Name = "column_create", Idempotent = true), Description("Create a typed list or library column. Arbitrary field XML is not accepted.")]
    public async Task<string> ColumnCreate(string listName, ColumnDefinition definition, CancellationToken cancellationToken) => Json(await site.CreateColumnAsync(listName, definition, cancellationToken));

    [McpServerTool(Name = "column_update", Idempotent = true), Description("Update a typed list or library column.")]
    public async Task<string> ColumnUpdate(string listName, string columnName, ColumnDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdateColumnAsync(listName, columnName, definition, cancellationToken));

    [McpServerTool(Name = "column_delete", Destructive = true, Idempotent = true), Description("Delete a column and its values; dryRun defaults true.")]
    public async Task<string> ColumnDelete(string listName, string columnName, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteColumnAsync(listName, columnName, dryRun, cancellationToken));

    [McpServerTool(Name = "item_get", ReadOnly = true), Description("Get one list item by positive numeric ID.")]
    public async Task<string> ItemGet(string listName, string itemId, CancellationToken cancellationToken) => Json(await site.GetItemAsync(listName, itemId, cancellationToken));

    [McpServerTool(Name = "item_query", ReadOnly = true), Description("Query list items with typed safe predicates; arbitrary OData is not accepted.")]
    public async Task<string> ItemQuery(string listName, McpItemQuery query, CancellationToken cancellationToken) => Json(await site.QueryItemsPageAsync(listName, query.ToCore(), cancellationToken));

    [McpServerTool(Name = "item_create"), Description("Create one list item from a bounded field-value object.")]
    public async Task<string> ItemCreate(string listName, JsonObject fields, CancellationToken cancellationToken) => Json(await site.CreateItemAsync(listName, fields, cancellationToken));

    [McpServerTool(Name = "item_update", Idempotent = true), Description("Update one list item from a bounded field-value object.")]
    public async Task<string> ItemUpdate(string listName, string itemId, JsonObject fields, string? expectedETag = null, CancellationToken cancellationToken = default) => Json(await site.UpdateItemAsync(listName, itemId, fields, expectedETag, cancellationToken));

    [McpServerTool(Name = "item_delete", Destructive = true, Idempotent = true), Description("Delete one list item; dryRun defaults true.")]
    public async Task<string> ItemDelete(string listName, string itemId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteItemAsync(listName, itemId, dryRun, cancellationToken));

    [McpServerTool(Name = "item_recycle", Destructive = true, Idempotent = true), Description("Move one list item to the configured site's recycle bin; dryRun defaults true.")]
    public async Task<string> ItemRecycle(string listName, string itemId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.RecycleItemAsync(listName, itemId, dryRun, cancellationToken));

    [McpServerTool(Name = "item_version_list", ReadOnly = true), Description("List recoverable versions for one configured-site list item.")]
    public async Task<string> ItemVersionList(string listName, string itemId, CancellationToken cancellationToken) => Json(await site.GetItemVersionsAsync(listName, itemId, cancellationToken));

    [McpServerTool(Name = "item_version_restore", Destructive = true, Idempotent = true), Description("Restore one list-item version; dryRun defaults true.")]
    public async Task<string> ItemVersionRestore(string listName, string itemId, string versionId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.RestoreItemVersionAsync(listName, itemId, versionId, dryRun, cancellationToken));

    [McpServerTool(Name = "view_list", ReadOnly = true), Description("List views for one configured-site list when the active backend supports SharePoint REST.")]
    public async Task<string> ViewList(string listName, CancellationToken cancellationToken) => Json(await site.ListViewsAsync(listName, cancellationToken));

    [McpServerTool(Name = "view_get", ReadOnly = true), Description("Get one typed view definition from a configured-site list.")]
    public async Task<string> ViewGet(string listName, string viewName, CancellationToken cancellationToken) => Json(await site.GetViewAsync(listName, viewName, cancellationToken));

    [McpServerTool(Name = "view_create", Idempotent = true), Description("Create a typed list view; arbitrary CAML is not accepted.")]
    public async Task<string> ViewCreate(string listName, McpViewDefinition definition, CancellationToken cancellationToken) => Json(await site.CreateViewAsync(listName, definition.ToCore(), cancellationToken));

    [McpServerTool(Name = "view_update", Idempotent = true), Description("Update a typed list view.")]
    public async Task<string> ViewUpdate(string listName, string viewName, McpViewDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdateViewAsync(listName, viewName, definition.ToCore(), cancellationToken));

    [McpServerTool(Name = "view_delete", Destructive = true, Idempotent = true), Description("Delete a list view; dryRun defaults true.")]
    public async Task<string> ViewDelete(string listName, string viewName, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteViewAsync(listName, viewName, dryRun, cancellationToken));

    [McpServerTool(Name = "view_set_default", Idempotent = true), Description("Set the default view for one list.")]
    public async Task<string> ViewSetDefault(string listName, string viewName, CancellationToken cancellationToken) => Json(await site.SetDefaultViewAsync(listName, viewName, cancellationToken));

    [McpServerTool(Name = "form_get", ReadOnly = true), Description("Get modern form layout for one configured-site list.")]
    public async Task<string> FormGet(string listName, CancellationToken cancellationToken) => Json(await site.GetFormAsync(listName, cancellationToken));

    [McpServerTool(Name = "configure_form", Idempotent = true), Description("Configure modern form header, sections/body, footer, field order, and read-only fields.")]
    public async Task<string> ConfigureForm(string listName, FormDefinition definition, CancellationToken cancellationToken) => Json(await site.ConfigureFormAsync(listName, definition, cancellationToken));

    [McpServerTool(Name = "field_visibility_get", ReadOnly = true), Description("Read New/Edit/Display visibility and conditional formulas independently from list form layout.")]
    public async Task<string> FieldVisibilityGet(string listName, CancellationToken cancellationToken) => Json(await site.GetFieldVisibilityRulesAsync(listName, cancellationToken));

    [McpServerTool(Name = "field_visibility_set", Idempotent = true), Description("Set a field's New/Edit/Display visibility and optional validated conditional-show expression.")]
    public async Task<string> FieldVisibilitySet(string listName, FieldVisibilityRule rule, CancellationToken cancellationToken) => Json(await site.SetFieldVisibilityRuleAsync(listName, rule, cancellationToken));

    [McpServerTool(Name = "form_format_set", Idempotent = true), Description("Set modern form formatting using a validated typed FormDefinition.")]
    public async Task<string> FormFormatSet(string listName, FormDefinition definition, CancellationToken cancellationToken) => Json(await site.ConfigureFormAsync(listName, definition, cancellationToken));

    [McpServerTool(Name = "column_format_set", Idempotent = true), Description("Set validated official SharePoint JSON column formatting.")]
    public async Task<string> ColumnFormatSet(string listName, string columnName, JsonObject formatting, CancellationToken cancellationToken) => Json(await site.SetColumnFormattingAsync(listName, columnName, formatting, cancellationToken));

    [McpServerTool(Name = "column_format_clear", Idempotent = true), Description("Clear column formatting.")]
    public async Task<string> ColumnFormatClear(string listName, string columnName, CancellationToken cancellationToken) => Json(await site.ClearColumnFormattingAsync(listName, columnName, cancellationToken));

    [McpServerTool(Name = "view_format_set", Idempotent = true), Description("Set validated official SharePoint JSON view formatting.")]
    public async Task<string> ViewFormatSet(string listName, string viewName, JsonObject formatting, CancellationToken cancellationToken) => Json(await site.SetViewFormattingAsync(listName, viewName, formatting, cancellationToken));

    [McpServerTool(Name = "view_format_clear", Idempotent = true), Description("Clear view formatting.")]
    public async Task<string> ViewFormatClear(string listName, string viewName, CancellationToken cancellationToken) => Json(await site.ClearViewFormattingAsync(listName, viewName, cancellationToken));

    [McpServerTool(Name = "library_list", ReadOnly = true), Description("List document libraries in the configured site.")]
    public async Task<string> LibraryList(CancellationToken cancellationToken) => Json(await site.ListListsAsync(ListKind.DocumentLibrary, cancellationToken));

    [McpServerTool(Name = "library_get", ReadOnly = true), Description("Get one document library by site-local title.")]
    public async Task<string> LibraryGet(string libraryName, CancellationToken cancellationToken) => Json(await site.GetListAsync(libraryName, cancellationToken));

    [McpServerTool(Name = "library_create", Idempotent = true), Description("Create a document library in the configured site.")]
    public async Task<string> LibraryCreate(ListDefinition definition, CancellationToken cancellationToken) => Json(await site.CreateListAsync(definition with { Kind = ListKind.DocumentLibrary }, cancellationToken));

    [McpServerTool(Name = "library_update", Idempotent = true), Description("Update Graph-supported document library properties.")]
    public async Task<string> LibraryUpdate(string libraryName, ListDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdateListAsync(libraryName, definition with { Kind = ListKind.DocumentLibrary }, cancellationToken));

    [McpServerTool(Name = "library_delete", Destructive = true, Idempotent = true), Description("Delete a document library and all content; dryRun defaults true.")]
    public async Task<string> LibraryDelete(string libraryName, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteListAsync(libraryName, dryRun, cancellationToken));

    [McpServerTool(Name = "file_list", ReadOnly = true), Description("List files/folders beneath a path whose first segment is a configured-site document library.")]
    public async Task<string> FileList(string relativePath, CancellationToken cancellationToken) => Json(await site.ListFilesAsync(relativePath, cancellationToken));

    [McpServerTool(Name = "file_get_metadata", ReadOnly = true), Description("Read metadata for one site-relative library path.")]
    public async Task<string> FileGetMetadata(string relativePath, CancellationToken cancellationToken) => Json(await site.GetFileMetadataAsync(relativePath, cancellationToken));

    [McpServerTool(Name = "file_read", ReadOnly = true), Description("Read a file as base64 without exposing a Microsoft download URL or token.")]
    public async Task<string> FileRead(string relativePath, CancellationToken cancellationToken)
    {
        var bytes = await site.ReadFileAsync(relativePath, cancellationToken);
        return Json(new { relativePath, length = bytes.LongLength, base64 = Convert.ToBase64String(bytes) });
    }

    [McpServerTool(Name = "file_write", Idempotent = true), Description("Write base64 file content to a validated site-relative library path.")]
    public async Task<string> FileWrite(string relativePath, string base64, bool overwrite = false, CancellationToken cancellationToken = default) => Json(await site.WriteFileAsync(relativePath, Convert.FromBase64String(base64), overwrite, cancellationToken));

    [McpServerTool(Name = "file_upload", Idempotent = true), Description("Upload base64 file content to a validated site-relative library path.")]
    public Task<string> FileUpload(string relativePath, string base64, bool overwrite = false, CancellationToken cancellationToken = default) => FileWrite(relativePath, base64, overwrite, cancellationToken);

    [McpServerTool(Name = "file_delete", Destructive = true, Idempotent = true), Description("Delete a file; dryRun defaults true.")]
    public async Task<string> FileDelete(string relativePath, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteFileAsync(relativePath, dryRun, cancellationToken));

    [McpServerTool(Name = "file_move", Idempotent = true), Description("Move a file between validated paths within the configured site.")]
    public async Task<string> FileMove(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken) => Json(await site.MoveFileAsync(sourceRelativePath, destinationRelativePath, cancellationToken));

    [McpServerTool(Name = "file_copy", Idempotent = true), Description("Copy a file between validated paths within the configured site.")]
    public async Task<string> FileCopy(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken) => Json(await site.CopyFileAsync(sourceRelativePath, destinationRelativePath, cancellationToken));

    [McpServerTool(Name = "file_copy_if_absent", Idempotent = true), Description("Copy a configured-site file only when the exact destination is absent; returns the existing destination instead of creating a suffixed duplicate.")]
    public async Task<string> FileCopyIfAbsent(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken) => Json(await site.CopyFileIfAbsentAsync(sourceRelativePath, destinationRelativePath, cancellationToken));

    [McpServerTool(Name = "file_recycle", Destructive = true, Idempotent = true), Description("Move a file or folder to the SharePoint recycle bin; dryRun defaults true.")]
    public async Task<string> FileRecycle(string relativePath, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.RecycleFileAsync(relativePath, dryRun, cancellationToken));

    [McpServerTool(Name = "file_version_list", ReadOnly = true), Description("List versions of an exact configured-site file path.")]
    public async Task<string> FileVersionList(string relativePath, CancellationToken cancellationToken) => Json(await site.GetFileVersionsAsync(relativePath, cancellationToken));

    [McpServerTool(Name = "file_version_restore", Destructive = true, Idempotent = true), Description("Restore a version of an exact configured-site file path; dryRun defaults true.")]
    public async Task<string> FileVersionRestore(string relativePath, string versionId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.RestoreFileVersionAsync(relativePath, versionId, dryRun, cancellationToken));

    [McpServerTool(Name = "folder_create", Idempotent = true), Description("Create a folder at a validated site-relative library path.")]
    public async Task<string> FolderCreate(string relativePath, CancellationToken cancellationToken) => Json(await site.CreateFolderAsync(relativePath, cancellationToken));

    [McpServerTool(Name = "folder_list", ReadOnly = true), Description("List one folder at a validated site-relative library path.")]
    public Task<string> FolderList(string relativePath, CancellationToken cancellationToken) => FileList(relativePath, cancellationToken);

    [McpServerTool(Name = "folder_delete", Destructive = true, Idempotent = true), Description("Delete a folder and descendants; dryRun defaults true.")]
    public async Task<string> FolderDelete(string relativePath, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteFolderAsync(relativePath, dryRun, cancellationToken));

    [McpServerTool(Name = "folder_move", Idempotent = true), Description("Move a folder within one configured-site library.")]
    public async Task<string> FolderMove(string sourceRelativePath, string destinationRelativePath, CancellationToken cancellationToken) => Json(await site.MoveFolderAsync(sourceRelativePath, destinationRelativePath, cancellationToken));

    [McpServerTool(Name = "evidence_search", ReadOnly = true), Description("Search file names and optionally locally extracted document text only in configured-site libraries; no tenant-wide or cross-site search.")]
    public async Task<string> EvidenceSearch(string query, int maxResults = 25, bool includeContent = false, CancellationToken cancellationToken = default) => Json(await site.SearchEvidenceAsync(query, maxResults, includeContent, cancellationToken));

    [McpServerTool(Name = "evidence_read_text", ReadOnly = true), Description("Read text extracted locally from a configured-site text, Word, PowerPoint, or Excel file without exposing a download URL or token.")]
    public async Task<string> EvidenceReadText(string relativePath, int maxCharacters = 200000, CancellationToken cancellationToken = default) => Json(await site.ReadExtractedTextAsync(relativePath, maxCharacters, cancellationToken));

    [McpServerTool(Name = "content_type_list", ReadOnly = true), Description("List site-local or list-local content types. Tenant content-type hub administration is excluded.")]
    public async Task<string> ContentTypeList(string? listName = null, CancellationToken cancellationToken = default) => Json(await site.ListContentTypesAsync(listName, cancellationToken));

    [McpServerTool(Name = "content_type_create", Idempotent = true), Description("Create a content type in the configured site only.")]
    public async Task<string> ContentTypeCreate(ContentTypeDefinition definition, CancellationToken cancellationToken) => Json(await site.CreateContentTypeAsync(definition, cancellationToken));

    [McpServerTool(Name = "content_type_update", Idempotent = true), Description("Update a content type in the configured site only.")]
    public async Task<string> ContentTypeUpdate(string contentTypeId, ContentTypeDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdateContentTypeAsync(contentTypeId, definition, cancellationToken));

    [McpServerTool(Name = "content_type_add_to_list", Idempotent = true), Description("Add a site content type to one configured-site list.")]
    public async Task<string> ContentTypeAddToList(string listName, string contentTypeId, CancellationToken cancellationToken) => Json(await site.AddContentTypeToListAsync(listName, contentTypeId, cancellationToken));

    [McpServerTool(Name = "page_list", ReadOnly = true), Description("List modern pages in the configured site.")]
    public async Task<string> PageList(CancellationToken cancellationToken) => Json(await site.ListPagesAsync(cancellationToken));

    [McpServerTool(Name = "page_get", ReadOnly = true), Description("Get one modern page and its canvas from the configured site.")]
    public async Task<string> PageGet(string pageName, CancellationToken cancellationToken) => Json(await site.GetPageAsync(pageName, cancellationToken));

    [McpServerTool(Name = "page_create", Idempotent = true), Description("Create a modern page using typed sections and friendly web-part definitions.")]
    public async Task<string> PageCreate(PageDefinition definition, CancellationToken cancellationToken) => Json(await site.CreatePageAsync(definition, cancellationToken));

    [McpServerTool(Name = "page_update", Idempotent = true), Description("Replace supported modern page properties and typed canvas layout.")]
    public async Task<string> PageUpdate(string pageName, PageDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdatePageAsync(pageName, definition, cancellationToken));

    [McpServerTool(Name = "page_delete", Destructive = true, Idempotent = true), Description("Delete a modern page; dryRun defaults true.")]
    public async Task<string> PageDelete(string pageName, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeletePageAsync(pageName, dryRun, cancellationToken));

    [McpServerTool(Name = "page_publish", Idempotent = true), Description("Publish a modern page.")]
    public async Task<string> PagePublish(string pageName, CancellationToken cancellationToken) => Json(await site.PublishPageAsync(pageName, cancellationToken));

    [McpServerTool(Name = "page_unpublish", Idempotent = true), Description("Unpublish a modern page when the active backend supports it.")]
    public async Task<string> PageUnpublish(string pageName, CancellationToken cancellationToken) => Json(await site.UnpublishPageAsync(pageName, cancellationToken));

    [McpServerTool(Name = "page_set_home", Idempotent = true), Description("Set a modern page as home when the active backend supports it.")]
    public async Task<string> PageSetHome(string pageName, CancellationToken cancellationToken) => Json(await site.SetHomePageAsync(pageName, cancellationToken));

    [McpServerTool(Name = "page_copy_if_absent", Idempotent = true), Description("Clone one configured-site modern page as a draft only when the exact destination page name is absent.")]
    public async Task<string> PageCopyIfAbsent(string sourcePageName, string destinationPageName, CancellationToken cancellationToken) => Json(await site.CopyPageIfAbsentAsync(sourcePageName, destinationPageName, cancellationToken));

    [McpServerTool(Name = "page_update_draft", Idempotent = true), Description("Replace the complete typed page canvas as a draft, optionally requiring the current ETag.")]
    public async Task<string> PageUpdateDraft(string pageName, PageDefinition definition, string? expectedETag = null, CancellationToken cancellationToken = default) => Json(await site.SavePageDraftAsync(pageName, definition, expectedETag, cancellationToken));

    [McpServerTool(Name = "page_version_list", ReadOnly = true), Description("List version history for one configured-site modern page.")]
    public async Task<string> PageVersionList(string pageName, CancellationToken cancellationToken) => Json(await site.GetPageVersionsAsync(pageName, cancellationToken));

    [McpServerTool(Name = "page_version_restore", Destructive = true, Idempotent = true), Description("Restore a page version as the current draft; dryRun defaults true.")]
    public async Task<string> PageVersionRestore(string pageName, string versionLabel, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.RestorePageVersionAsync(pageName, versionLabel, dryRun, cancellationToken));

    [McpServerTool(Name = "page_recycle", Destructive = true, Idempotent = true), Description("Move a configured-site page to the SharePoint recycle bin; dryRun defaults true.")]
    public async Task<string> PageRecycle(string pageName, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.RecyclePageAsync(pageName, dryRun, cancellationToken));

    [McpServerTool(Name = "page_section_add", Idempotent = true), Description("Add a section by applying a complete typed page definition; Graph reconciles the canvas atomically.")]
    public Task<string> PageSectionAdd(string pageName, PageDefinition completePageDefinition, CancellationToken cancellationToken) => PageUpdate(pageName, completePageDefinition, cancellationToken);

    [McpServerTool(Name = "page_section_update", Idempotent = true), Description("Update a section by applying a complete typed page definition; Graph reconciles the canvas atomically.")]
    public Task<string> PageSectionUpdate(string pageName, PageDefinition completePageDefinition, CancellationToken cancellationToken) => PageUpdate(pageName, completePageDefinition, cancellationToken);

    [McpServerTool(Name = "page_section_delete", Destructive = true, Idempotent = true), Description("Delete a section by applying a complete typed page definition with dryRun support.")]
    public async Task<string> PageSectionDelete(string pageName, PageDefinition completePageDefinition, bool dryRun = true, CancellationToken cancellationToken = default) => dryRun
        ? Json(new MutationResult("page_section_delete", pageName, false, "Dry run: would reconcile the page without the omitted section."))
        : await PageUpdate(pageName, completePageDefinition, cancellationToken);

    [McpServerTool(Name = "webpart_available", ReadOnly = true), Description("List friendly web-part types actually available through the configured backend and their support status.")]
    public async Task<string> WebPartAvailable(CancellationToken cancellationToken) => Json(await site.GetAvailableWebPartsAsync(cancellationToken));

    [McpServerTool(Name = "webpart_get", ReadOnly = true), Description("Get a previously observed web part by ID from a configured-site page.")]
    public async Task<string> WebPartGet(string pageName, string webPartId, CancellationToken cancellationToken) => Json(await site.GetWebPartAsync(pageName, webPartId, cancellationToken));

    [McpServerTool(Name = "webpart_add", Idempotent = true), Description("Add a supported friendly web-part definition to a page section.")]
    public async Task<string> WebPartAdd(string pageName, string sectionKey, WebPartDefinition definition, CancellationToken cancellationToken) => Json(await site.AddWebPartAsync(pageName, sectionKey, definition, cancellationToken));

    [McpServerTool(Name = "webpart_update", Idempotent = true), Description("Update a supported web part on one configured-site page.")]
    public async Task<string> WebPartUpdate(string pageName, string webPartId, WebPartDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdateWebPartAsync(pageName, webPartId, definition, cancellationToken));

    [McpServerTool(Name = "webpart_move", Idempotent = true), Description("Move a web part when the active backend supports reliable standalone movement.")]
    public async Task<string> WebPartMove(string pageName, string webPartId, string sectionKey, int column, int order, CancellationToken cancellationToken) => Json(await site.MoveWebPartAsync(pageName, webPartId, sectionKey, column, order, cancellationToken));

    [McpServerTool(Name = "webpart_delete", Destructive = true, Idempotent = true), Description("Delete a web part; dryRun defaults true.")]
    public async Task<string> WebPartDelete(string pageName, string webPartId, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.DeleteWebPartAsync(pageName, webPartId, dryRun, cancellationToken));

    [McpServerTool(Name = "navigation_get", ReadOnly = true), Description("Get configured-site navigation when the active backend supports SharePoint REST.")]
    public async Task<string> NavigationGet(NavigationLocation location = NavigationLocation.QuickLaunch, CancellationToken cancellationToken = default) => Json(await site.GetNavigationAsync(location, cancellationToken));

    [McpServerTool(Name = "navigation_add", Idempotent = true), Description("Add a site-local navigation entry.")]
    public async Task<string> NavigationAdd(NavigationDefinition definition, CancellationToken cancellationToken) => Json(await site.AddNavigationAsync(definition, cancellationToken));

    [McpServerTool(Name = "navigation_update", Idempotent = true), Description("Update a keyed site-local navigation entry.")]
    public async Task<string> NavigationUpdate(string key, NavigationDefinition definition, CancellationToken cancellationToken) => Json(await site.UpdateNavigationAsync(key, definition, cancellationToken));

    [McpServerTool(Name = "navigation_remove", Destructive = true, Idempotent = true), Description("Remove a keyed site-local navigation entry; dryRun defaults true.")]
    public async Task<string> NavigationRemove(string key, NavigationLocation location = NavigationLocation.QuickLaunch, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await site.RemoveNavigationAsync(key, location, dryRun, cancellationToken));

    [McpServerTool(Name = "navigation_reorder", Idempotent = true), Description("Reorder configured-site navigation using stable keys.")]
    public async Task<string> NavigationReorder(IReadOnlyList<string> keys, NavigationLocation location = NavigationLocation.QuickLaunch, string? parentKey = null, CancellationToken cancellationToken = default) => Json(await site.ReorderNavigationAsync(keys, location, parentKey, cancellationToken));

    [McpServerTool(Name = "permissions_inspect", ReadOnly = true), Description("Inspect permissions only on the configured site or a validated site-relative resource.")]
    public async Task<string> PermissionsInspect(PermissionTarget? target = null, CancellationToken cancellationToken = default) => Json(await site.InspectPermissionsAsync(target, cancellationToken));

    [McpServerTool(Name = "permissions_grant"), Description("Grant a site-local permission only when explicitly enabled in trusted configuration and supported by the backend.")]
    public async Task<string> PermissionsGrant(string principal, string role, PermissionTarget? target = null, CancellationToken cancellationToken = default) => Json(await site.GrantPermissionAsync(principal, role, target, cancellationToken));

    [McpServerTool(Name = "permissions_revoke", Destructive = true), Description("Revoke a site-local permission only when explicitly enabled in trusted configuration.")]
    public async Task<string> PermissionsRevoke(string principal, string role, PermissionTarget? target = null, CancellationToken cancellationToken = default) => Json(await site.RevokePermissionAsync(principal, role, target, cancellationToken));

    [McpServerTool(Name = "permissions_reset_inheritance", Destructive = true), Description("Reset inheritance only for a validated configured-site resource when explicitly enabled.")]
    public async Task<string> PermissionsResetInheritance(PermissionTarget target, CancellationToken cancellationToken) => Json(await site.ResetPermissionInheritanceAsync(target, cancellationToken));

    [McpServerTool(Name = "site_validate_blueprint", ReadOnly = true), Description("Parse and validate a JSON or YAML-equivalent site blueprint without calling Microsoft APIs.")]
    public string SiteValidateBlueprint(string document) => Json(blueprintValidator.Validate(BlueprintParser.Parse(document)));

    [McpServerTool(Name = "site_plan_blueprint", ReadOnly = true), Description("Inspect the configured site and plan a validated blueprint without mutation.")]
    public async Task<string> SitePlanBlueprint(string document, CancellationToken cancellationToken) => Json(blueprints.Plan(BlueprintParser.Parse(document), await site.GetInventoryAsync(cancellationToken)));

    [McpServerTool(Name = "site_diff_blueprint", ReadOnly = true), Description("Return the current configured-site diff for a validated blueprint without mutation.")]
    public Task<string> SiteDiffBlueprint(string document, CancellationToken cancellationToken) => SitePlanBlueprint(document, cancellationToken);

    [McpServerTool(Name = "site_apply_blueprint", Idempotent = true), Description("Validate, diff, apply, read back, and verify a JSON/YAML blueprint. dryRun defaults true.")]
    public async Task<string> SiteApplyBlueprint(string document, bool dryRun = true, CancellationToken cancellationToken = default) => Json(await blueprints.ApplyAsync(BlueprintParser.Parse(document), dryRun, cancellationToken));

    [McpServerTool(Name = "preview_change", ReadOnly = true), Description("Preview a complete declarative SharePoint blueprint against current state without mutation.")]
    public Task<string> PreviewChange(string document, CancellationToken cancellationToken) => SitePlanBlueprint(document, cancellationToken);

    [McpServerTool(Name = "snapshot_resource", ReadOnly = true), Description("Capture fresh site-bound state and available versions for one typed resource before a change.")]
    public async Task<string> SnapshotResource(SnapshotResourceKind kind, string? resourceName = null, CancellationToken cancellationToken = default)
    {
        var snapshot = kind switch
        {
            SnapshotResourceKind.Site => new JsonObject
            {
                ["info"] = JsonSerializer.SerializeToNode(await site.GetInfoAsync(cancellationToken), JsonDefaults.Options),
                ["appearance"] = await site.GetSiteAppearanceAsync(cancellationToken),
                ["theme"] = await site.GetThemeAsync(cancellationToken),
                ["quickLaunch"] = await site.GetNavigationAsync(NavigationLocation.QuickLaunch, cancellationToken),
                ["topNavigation"] = await site.GetNavigationAsync(NavigationLocation.TopNavigationBar, cancellationToken),
                ["footer"] = await site.GetNavigationAsync(NavigationLocation.Footer, cancellationToken)
            },
            SnapshotResourceKind.List or SnapshotResourceKind.Library => await SnapshotList(resourceName, cancellationToken),
            SnapshotResourceKind.Page => new JsonObject
            {
                ["page"] = await site.GetPageAsync(RequireResourceName(resourceName), cancellationToken),
                ["versions"] = await site.GetPageVersionsAsync(RequireResourceName(resourceName), cancellationToken)
            },
            SnapshotResourceKind.File => new JsonObject
            {
                ["metadata"] = await site.GetFileMetadataAsync(RequireResourceName(resourceName), cancellationToken),
                ["versions"] = await site.GetFileVersionsAsync(RequireResourceName(resourceName), cancellationToken)
            },
            _ => throw new BoundaryViolationException("Unsupported snapshot resource kind.")
        };
        return Json(new { kind, resourceName, capturedAt = DateTimeOffset.UtcNow, snapshot });
    }

    private async Task<JsonObject> SnapshotList(string? resourceName, CancellationToken cancellationToken)
    {
        var name = RequireResourceName(resourceName);
        var result = new JsonObject
        {
            ["list"] = await site.GetListAsync(name, cancellationToken),
            ["columns"] = await site.ListColumnsAsync(name, cancellationToken),
            ["views"] = await site.ListViewsAsync(name, cancellationToken)
        };
        try { result["form"] = await site.GetFormAsync(name, cancellationToken); }
        catch (UnsupportedCapabilityException exception) { result["formUnavailable"] = exception.Message; }
        return result;
    }

    private static string RequireResourceName(string? value) => string.IsNullOrWhiteSpace(value)
        ? throw new BoundaryViolationException("resourceName is required for this snapshot kind.")
        : value;

    private static string Json<T>(T value) => JsonSerializer.Serialize(value, JsonDefaults.Options);
}
