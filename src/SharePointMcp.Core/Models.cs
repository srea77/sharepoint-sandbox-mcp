using System.Text.Json.Nodes;

namespace SharePointMcp.Core;

public enum SupportLevel { Complete, Partial, Experimental, BlockedByAuthentication, BlockedByApi, RequiresHumanOrAdmin, NotImplemented }
public enum FieldKind { Text, MultilineText, Number, Currency, DateTime, Boolean, Choice, MultiChoice, Person, Lookup, Hyperlink, Calculated, Taxonomy }
public enum ListKind { GenericList, DocumentLibrary }
public enum ChangeKind { Create, Update, Delete, NoOp, Unsupported }
public enum RiskLevel { None, Low, Destructive }
public enum NavigationLocation { QuickLaunch, TopNavigationBar, Footer }
public enum PermissionResourceKind { Site, List, Library, Page }
public enum SnapshotResourceKind { Site, List, Library, Page, File }

public sealed record CapabilityStatus(string Family, SupportLevel Status, string Backend, string Notes);
public sealed record SiteInfo(string Id, string DisplayName, string WebUrl, string? Description);
public sealed record SiteIdentity(
    string? Id,
    string DisplayName,
    string? Email,
    string LoginName,
    JsonObject EffectivePermissions);
public sealed record SiteAppearanceUpdate(
    string? Title = null,
    string? Description = null,
    string? LogoSiteRelativePath = null,
    bool? QuickLaunchEnabled = null,
    string? HeaderLayout = null,
    int? HeaderEmphasis = null,
    bool? FooterEnabled = null,
    bool? MegaMenuEnabled = null,
    bool? NavigationAudienceTargetingEnabled = null);
public sealed record HealthReport(bool ConfigurationValid, bool AuthenticationReady, bool SiteReachable, IReadOnlyList<string> Checks);
public sealed record SiteInventory(
    JsonNode Site,
    JsonArray Lists,
    JsonArray Libraries,
    JsonArray Pages,
    JsonArray Navigation,
    IReadOnlyList<string> Warnings,
    JsonArray? ListDetails = null,
    JsonArray? PageDetails = null);
public sealed record MutationResult(string Operation, string Target, bool Changed, string Outcome, string? CorrelationId = null, JsonNode? Data = null);

public sealed record ListDefinition(
    string Title,
    string? Description = null,
    ListKind Kind = ListKind.GenericList,
    bool? EnableVersioning = null,
    bool? EnableAttachments = null,
    bool? Hidden = null);

public sealed record ColumnDefinition(
    string Name,
    string? DisplayName,
    FieldKind Type,
    bool Required = false,
    string? Description = null,
    JsonNode? DefaultValue = null,
    IReadOnlyList<string>? Choices = null,
    string? LookupList = null,
    string? LookupColumn = null,
    string? Formula = null,
    string? ValidationFormula = null,
    string? ValidationMessage = null,
    bool ReadOnly = false,
    bool Hidden = false,
    bool AllowMultipleValues = false);

public sealed record ItemQuery(
    IReadOnlyList<QueryPredicate>? Predicates = null,
    IReadOnlyList<QuerySort>? OrderBy = null,
    int Limit = 100,
    string? Cursor = null,
    IReadOnlyList<string>? SelectedFields = null,
    bool ExpandLookupValues = true);
public sealed record QueryPredicate(string Field, string Operator, JsonNode? Value);
public sealed record QuerySort(string Field, bool Descending = false);
public sealed record ItemQueryResult(JsonArray Items, string? NextCursor);

public sealed record ViewDefinition(
    string Name,
    IReadOnlyList<string> Columns,
    IReadOnlyList<QueryPredicate>? Filter = null,
    IReadOnlyList<QuerySort>? Sort = null,
    string? GroupBy = null,
    int RowLimit = 100,
    bool IsDefault = false);

public sealed record FormDefinition(
    JsonNode? Header = null,
    JsonNode? Body = null,
    JsonNode? Footer = null,
    IReadOnlyList<FormSection>? Sections = null,
    IReadOnlyList<string>? FieldOrder = null,
    IReadOnlyList<string>? ReadOnlyFields = null);
public sealed record FormSection(string DisplayName, IReadOnlyList<string> Fields);
public sealed record FieldVisibilityRule(
    string FieldName,
    bool? ShowInNewForm = null,
    bool? ShowInEditForm = null,
    bool? ShowInDisplayForm = null,
    string? ConditionalShowExpression = null);
public sealed record PermissionTarget(PermissionResourceKind Kind = PermissionResourceKind.Site, string? Name = null);

public sealed record PageDefinition(
    string Name,
    string Title,
    string? Description = null,
    IReadOnlyList<PageSectionDefinition>? Sections = null,
    bool Publish = false,
    bool SetAsHome = false,
    PageTitleRegionDefinition? TitleRegion = null);
public sealed record PageTitleRegionDefinition(
    string Layout = "plain",
    string TextAlignment = "left",
    string? ImageSitePath = null,
    string? AlternativeText = null,
    bool EnableGradientEffect = true,
    bool ShowAuthor = false,
    bool ShowPublishedDate = false,
    string? TextAboveTitle = null);
public sealed record PageSectionDefinition(
    string Key,
    string Layout,
    IReadOnlyList<WebPartDefinition> Components,
    string Emphasis = "none");
public sealed record WebPartDefinition(
    string Key,
    string Type,
    JsonObject Properties,
    int Column = 1,
    int Order = 1,
    HeroWebPartOptions? Hero = null,
    ListWebPartOptions? List = null,
    NewsWebPartOptions? News = null,
    EventsWebPartOptions? Events = null,
    FormsWebPartOptions? Forms = null,
    PowerAppsWebPartOptions? PowerApps = null,
    EmbedWebPartOptions? Embed = null);

public sealed record HeroWebPartOptions(string Layout, IReadOnlyList<HeroItemDefinition> Items);
public sealed record HeroItemDefinition(
    string Title,
    string? Description = null,
    string? Target = null,
    string? TargetList = null,
    string? ImageSitePath = null,
    string? AlternativeText = null,
    string CallToAction = "Open");
public sealed record ListWebPartOptions(string ListName, string? ViewName = null, bool ShowCommandBar = true, bool ShowSeeAll = true);
public sealed record NewsWebPartOptions(string Layout = "topStory", int MaxItems = 4, bool ShowTitle = true);
public sealed record EventsWebPartOptions(string? ListName = null, string Layout = "filmstrip", int MaxItems = 6, bool ShowTitle = true);
public sealed record FormsWebPartOptions(string FormId, string Mode = "collectResponses", string? Title = null);
public sealed record PowerAppsWebPartOptions(Guid AppId, bool ScaleToFit = true);
public sealed record EmbedWebPartOptions(Uri Url, bool ScaleToFit = true);

public sealed record SiteThemeDefinition(string Name, bool IsInverted, IReadOnlyDictionary<string, string> Palette);
public sealed record SiteThemeValidationResult(bool Valid, SiteThemeDefinition? Normalized, IReadOnlyList<string> Errors);

public sealed record NavigationDefinition(
    string Key,
    string Title,
    string? Target = null,
    string? TargetList = null,
    IReadOnlyList<NavigationDefinition>? Children = null,
    NavigationLocation Location = NavigationLocation.QuickLaunch,
    string? ParentKey = null,
    bool OpenInNewWindow = false,
    IReadOnlyList<Guid>? AudienceIds = null);
public sealed record ContentTypeDefinition(string Name, string? Description = null, string? ParentId = null, IReadOnlyList<string>? Columns = null);

public sealed record BlueprintChange(ChangeKind Kind, string ResourceType, string ResourceKey, RiskLevel Risk, string Summary);
public sealed record BlueprintReport(bool Valid, bool DryRun, IReadOnlyList<BlueprintChange> Changes, IReadOnlyList<string> Errors, IReadOnlyList<string> Warnings)
{
    public bool HasChanges => Changes.Any(x => x.Kind is not ChangeKind.NoOp);
}
