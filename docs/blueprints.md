# Blueprint schema

Blueprints are JSON or YAML documents with `schemaVersion: 1`. The canonical example is [equipment-portal.blueprint.json](../samples/equipment-portal.blueprint.json).

## Top level

```yaml
schemaVersion: 1
site:
  title: Equipment Request Portal
  description: Request, approve, and document employee equipment.
  theme:
    name: Equipment Portal
    isInverted: false
    palette:
      themePrimary: "#005A9E"
      neutralPrimary: "#242424"
      white: "#FFFFFF"
lists: []
pages: []
navigation: []
```

`site.title` and `site.description` are descriptive. `site.theme` is operational: the palette accepts only documented semantic slot names and canonical six-digit hex colors, requires `themePrimary`, `neutralPrimary`, and `white`, and is applied only through the configured site's exact `ApplyTheme` endpoint. CSS, URLs, fonts, and tenant-theme administration are rejected.

## Lists and libraries

Each list has `name`, optional `description`, `kind` (`genericList` or `documentLibrary`), optional `enableVersioning`/`enableAttachments`, `columns`, `views`, optional `form`, `columnFormatting`, and `viewFormatting`.

Column types are `text`, `multilineText`, `number`, `currency`, `dateTime`, `boolean`, `choice`, `multiChoice`, `person`, `lookup`, `hyperlink`, `calculated`, and `taxonomy`. Choice types require non-empty `choices`; lookup requires `lookupList` and `lookupColumn`; calculated requires `formula`. Taxonomy validates with a warning but is blocked by the Graph backend.

Views use typed `columns`, `filter`, `sort`, optional `groupBy`, `rowLimit` from 1 to 5000, and `isDefault`. No raw CAML or OData is accepted.

Form sections use `displayName` plus field-name arrays. Header/body/footer and column/view formatting can carry official SharePoint JSON. JSON is bounded to 256 KB and rejects active `<script` or `javascript:` content before any backend call.

## Pages

Page fields are `name`, `title`, optional `description`, `titleRegion`, `sections`, `publish`, and `setAsHome`. Title regions support `imageAndTitle`, `plain`, `colorBlock`, and `overlap`, left/center alignment, optional site-relative image plus required alt text, gradient, author/date visibility, and text above the title. Supported section layouts are `oneColumn`, `twoColumn`, `threeColumn`, `oneThirdLeft`, `oneThirdRight`, `fullWidth`, and at most one `verticalSection`; emphasis is `none`, `neutral`, `soft`, or `strong`.

Components have stable `key`, friendly `type`, `properties`, and positive `column`/`order`. The schema recognizes text, image, hero, quick links, list, document library, people, news, events, button, divider, spacer, embed, Forms, and Power Apps. `webpart_available` is authoritative for what the active backend can actually persist.

Advanced types require one matching typed options object and an empty `properties` object. Raw PnP/component property JSON is rejected. Lists, libraries, views, and Hero list targets must be declared in the same blueprint and are resolved against the configured site by the trusted backend:

```yaml
- key: equipment-hero
  type: hero
  properties: {}
  hero:
    layout: tiles
    items:
      - title: Submit a request
        targetList: Equipment Requests
        callToAction: Request equipment
- key: pending
  type: list
  properties: {}
  list:
    listName: Equipment Requests
    viewName: Pending Approval
    showCommandBar: true
    showSeeAll: true
- key: docs
  type: documentLibrary
  properties: {}
  list:
    listName: Equipment Documentation
- key: news
  type: news
  properties: {}
  news: { layout: list, maxItems: 4, showTitle: true }
- key: events
  type: events
  properties: {}
  events: { layout: filmstrip, maxItems: 6, showTitle: true }
```

Hero currently supports one to five `tiles`; SharePoint responsively presents them as a carousel on a narrow screen. Each item uses exactly one site-relative `target` or declared `targetList`. A custom `imageSitePath` must be site-relative and requires `alternativeText`.

Microsoft Forms and Power Apps are experimental typed embeds. They accept only a provider form identifier/mode or an app GUID/scale flag; the MCP does not create or administer the external form/app. General Embed accepts only an HTTPS `url` whose exact FQDN is present in trusted `allowedExternalEmbedHosts`. It never accepts iframe HTML and cannot modify the allowlist or SharePoint HTML Field Security.

Arbitrary modern-page CSS, selectors, DOM mutation, and JavaScript are not blueprint capabilities. Bespoke visuals require a separately built, reviewed, and human/admin-installed SPFx component with component-scoped styles.

## Navigation

Entries have stable `key`, `title`, exactly one of relative `target` or local `targetList`, and optional `children`. The composite backend persists Quick Launch entries through site-bound REST; external URLs and cross-site targets are rejected.

## Operations

- `site_validate_blueprint`: parse and validate only.
- `site_plan_blueprint` / `site_diff_blueprint`: inspect the configured site and return intended create/update operations without mutation.
- `site_apply_blueprint`: defaults to `dryRun=true`; with false, applies supported operations, records unsupported operations, then reads inventory back.

Idempotency keys are case-insensitive list/library names, column internal/display names, view names, page names, navigation keys/titles, section keys, and component keys. Advanced controls receive deterministic instance IDs. A trusted definition hash plus live PnP page fingerprint makes an unchanged advanced page a true second-apply no-op. Reapplying avoids duplicates; unrelated Graph-backed resources may still report updates where the API lacks a reliable normalized equality comparison.
