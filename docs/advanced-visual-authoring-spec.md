# Advanced visual authoring specification

Status: Phases 0, 1, and 2 live verified on `feature/advanced-visual-authoring`; Phase 3 typed implementation complete but rendered acceptance is resource/configuration dependent
Baseline: `checkpoint/site-bound-v1` (`87d8a71`)
Scope: richer modern SharePoint pages without weakening the immutable site capability boundary

## 1. Objective

Extend the existing page, web-part, and blueprint models so an MCP client can produce polished modern SharePoint experiences using the components actually available in the configured sandbox site.

The first implementation slice targets:

- Hero;
- List and Document Library;
- News;
- Events;
- Microsoft Forms;
- Power Apps;
- bounded Embed;
- site theme application;
- richer title regions, full-width/vertical sections, and section emphasis.

The result must remain idempotent, typed, auditable, and impossible for the MCP caller to redirect to another SharePoint site or generic Microsoft resource.

## 2. Supported-platform findings

### Microsoft Graph

Graph officially creates text parts and only fourteen standard web-part types: Bing Maps, Button, Call to Action, Divider, Document Embed, Image, Image Gallery, Link Preview, Org Chart, People, Quick Links, Spacer, YouTube, and Title Area. Hero, List, Document Library, News, Events, Forms, Power Apps, and general Embed are not on that supported list and must not be forced through Graph.

Source: <https://learn.microsoft.com/graph/api/sitepage-create-webpart>

### PnP Core SDK

PnP Core supports modern pages, can enumerate the page components actually installed in a site with `AvailablePageComponentsAsync`, and exposes built-in component identities including Hero, List, News, Events, Microsoft Forms, Power Apps, and Content Embed. It accepts a custom `IAuthenticationProvider`, allowing the existing Azure Identity credential and encrypted cache to remain the authentication source.

Sources:

- <https://pnp.github.io/pnpcore/using-the-sdk/pages-webparts.html>
- <https://pnp.github.io/pnpcore/api/PnP.Core.Model.SharePoint.DefaultWebPart.html>
- <https://pnp.github.io/pnpcore/using-the-sdk/custom-authentication-provider.html>

PnP is a community-supported SDK rather than a Microsoft Graph contract. Every advanced component therefore needs a live capability probe and rendered acceptance test. Unsupported or changed property schemas must produce an explicit experimental/blocked result rather than a false success.

### Themes and CSS

SharePoint exposes an official site-scoped `/_api/thememanager/ApplyTheme` operation. Tenant theme creation/update/deletion is out of scope because it changes tenant-wide state. The MCP may apply a validated palette to the configured site but may not administer the tenant theme catalog.

Source: <https://learn.microsoft.com/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-rest-api>

Arbitrary CSS against modern SharePoint pages is not a supported extension mechanism. The SharePoint DOM and Microsoft CSS classes are not APIs. The supported custom-code path is SharePoint Framework (SPFx), whose styles must remain scoped to a trusted installed component.

Sources:

- <https://learn.microsoft.com/sharepoint/dev/solution-guidance/modern-experience-customizations-customize-pages>
- <https://learn.microsoft.com/sharepoint/dev/spfx/extensions/overview-extensions>

## 3. Architecture

```text
MCP tools / blueprint engine
          |
          v
ISandboxSite + typed visual models
          |
          +-- GraphPageComposer
          |     supported Graph text/standard parts
          |
          +-- PnPPageComposer
          |     Hero/List/Library/News/Events/Forms/PowerApps/Embed
          |
          +-- SiteThemeService
                exact ApplyTheme endpoint only
          |
          v
immutable SiteBoundary + fixed configured site
```

`PnPContext`, authentication providers, raw component manifests, raw property JSON write inputs, and generic REST/CSOM clients remain infrastructure-only. MCP tools continue to depend on `ISandboxSite` or narrower site-bound services. Standard Graph page readback may contain site-returned component metadata for inspection, but the caller cannot feed it back as an advanced write payload.

### Bound PnP authentication

Add a `BoundPnPAuthenticationProvider` that:

1. uses the existing `DeviceCodeCredential` and encrypted cache;
2. permits token acquisition only for the configured SharePoint hostname and the explicitly required Graph audience;
3. rejects any other resource/audience;
4. creates a `PnPContext` only from `TrustedSiteConfiguration.SiteUrl`;
5. is never serialized, returned, or injected into MCP tools.

The initial spike must record every outbound host/resource used by PnP. Any unexpected tenant-admin, cross-site, or directory call fails the spike.

## 4. Typed models

Replace free-form interpretation of advanced `properties` with validated discriminated definitions. The backend may still serialize component property JSON internally.

### Hero

```yaml
type: hero
layout: tiles | layers | carousel
autoplay: false
items:
  - key: submit
    title: Submit a request
    description: Request approved equipment
    target:
      list: Equipment Requests
    image:
      sitePath: SiteAssets/equipment-request.jpg
    callToAction: Start request
    altText: Laptop and monitor on a desk
    focalPoint: { x: 0.5, y: 0.45 }
```

Rules: one to five items; stable item keys; images resolve through a configured-site library; targets default to site-local pages/lists/files; external targets require a trusted-host allowlist.

### List and Document Library

```yaml
type: list
list: Equipment Requests
view: Pending Approval
title: Requests awaiting review
showCommandBar: true
showSeeAll: true
```

```yaml
type: documentLibrary
library: Equipment Documentation
view: Published Documents
folder: Policies
title: Equipment policies
```

The backend resolves list, library, view, and folder identifiers itself. Callers cannot submit GUIDs, server-relative site paths, or component property JSON.

### News

```yaml
type: news
source: currentSite
layout: topStory | list | tiles | carousel
maxItems: 5
showTitle: true
```

Only `currentSite` is supported initially. Hub, selected-site, and all-sites sources are rejected because they would make the rendered component a cross-site discovery surface.

### Events

```yaml
type: events
source: currentSite
list: Events
layout: compact | filmstrip
maxItems: 5
dateRange: upcoming
```

Only a list resolved inside the configured site may be used.

### Microsoft Forms

```yaml
type: microsoftForms
formId: validated-provider-id
mode: collectResponses | showResults
```

Form identifiers are data properties, never backend route identifiers. If a URL form is accepted, its host must be an approved Microsoft Forms host and HTTPS is mandatory.

### Power Apps

```yaml
type: powerApps
appId: 00000000-0000-0000-0000-000000000000
height: 600
```

An app GUID is accepted only as a typed property for the Power Apps component; it never enters a Graph/REST path and is never registered as a SharePoint resource ID. Availability and user authorization are verified by rendered inspection; the MCP does not create or administer the Power App.

### Embed

```yaml
type: embed
url: https://approved.example/resource
title: External status
height: 480
```

This means a bounded modern Embed web part, not arbitrary HTML or JavaScript. Requirements:

- HTTPS only;
- hostname must match a human-managed trusted allowlist in configuration;
- reject credentials, fragments containing script, `data:`, `javascript:`, `file:`, localhost, IP literals, and encoded host tricks;
- never modify SharePoint HTML Field Security or its domain allowlist through MCP;
- never accept raw iframe markup;
- readback must preserve the validated destination host.

## 5. Site theme model

Add coherent tools:

- `site_theme_get` — inspect the configured site's current theme identity where supported;
- `site_theme_validate` — validate a palette without mutation;
- `site_theme_apply` — apply a validated palette to this site only, with `dryRun=true` by default.

The model accepts named semantic color slots containing canonical six-digit hex colors and an optional inverted flag. It rejects CSS, URLs, fonts, scripts, and unknown keys. The transport allowlist gains only the exact `thememanager/ApplyTheme` endpoint. It must continue to reject `AddTenantTheme`, `UpdateTenantTheme`, `DeleteTenantTheme`, and every other `thememanager` operation.

Blueprint addition:

```yaml
site:
  title: Equipment Portal
  theme:
    name: Equipment Portal Theme
    isInverted: false
    palette:
      themePrimary: "#005A9E"
      themeDark: "#004578"
      themeLighterAlt: "#F3F9FD"
      neutralPrimary: "#242424"
      white: "#FFFFFF"
```

The diff engine normalizes color casing and reports a no-op when the effective theme already matches.

## 6. Richer page layout

Extend `PageDefinition` and `PageSectionDefinition` with:

- typed title-region layout, image, alt text, and focal point;
- section emphasis: `none`, `neutral`, `soft`, `strong`;
- full-width sections on communication sites only;
- vertical sections where the site/page supports them;
- collapsible section options only if the live component model exposes them reliably;
- explicit accessible heading hierarchy and image alternative text validation.

Unsupported combinations fail validation before mutation.

## 7. SPFx and “pixel-level CSS”

Raw site CSS, arbitrary selectors, DOM mutation, custom actions, and arbitrary JavaScript remain prohibited.

If bespoke visuals beyond native themes/web parts are still required, treat them as a separate optional deliverable:

1. build a versioned SPFx web part with Fluent UI and component-scoped styles;
2. security-review and package it outside MCP runtime;
3. require explicit human/admin approval for app-catalog deployment;
4. expose only typed properties for that already-installed component through MCP;
5. never expose SPFx package upload, arbitrary JavaScript, CSS, or tenant app-catalog administration as normal MCP tools.

This can provide precise layout inside the component's own rectangle, not unrestricted restyling of SharePoint chrome or page DOM.

## 8. Tool behavior

Retain the existing coherent surface:

- `webpart_available` returns live-discovered components, backend, support level, required fields, and experimental status;
- `page_create`, `page_update`, section tools, and blueprint apply accept the typed definitions above and reconcile a complete canvas;
- standalone `webpart_add`/`webpart_update` remain Graph-only because safely editing one advanced part would require reconstructing the complete typed page from an undocumented live manifest;
- blueprint plan/diff/apply handles the new definitions idempotently;
- theme operations use the three new site-theme tools.

Do not add generic component-ID, raw `PropertiesJson`, CanvasContent1, CSOM, PnP, REST, or script tools.

## 9. Delivery phases

### Phase 0 — capability spike

- Add PnP Core behind an infrastructure-only experimental adapter.
- Reuse the current encrypted credential through a bounded custom auth provider.
- Enumerate available components from the live sandbox and cache by site/web ID.
- Create a disposable unpublished page containing one empty instance of each target component.
- Read it back through PnP/Graph and visually inspect it.
- Record permissions, outbound hosts, component IDs, property schema versions, and failures.
- Delete the disposable page through the typed site abstraction after evidence is captured.

Exit criterion: at least Hero and List render using the current no-custom-app session without an outbound boundary violation.

Phase 0 result (2026-09-28): **passed** with PnP Core 1.18.0 and the existing cached delegated session. The trusted probe discovered Hero, List/Document Library, News, Events, Microsoft Forms, Power Apps, and Embed in the configured site. It created an unpublished disposable page containing Hero and List controls, read back both controls, then deleted the page. The bounded authentication provider observed only the configured SharePoint hostname; no tenant-admin or foreign SharePoint request was needed. The probe mutations were written to the trusted audit log.

### Phase 1 — native visual foundations

- Add Graph-supported Call to Action, Bing Maps, and Title Area mappings currently missing from the friendly map.
- Add title-region and section-emphasis models.
- Implement validated site theme apply/dry-run/audit/readback.
- Add unit, boundary, blueprint, and live rendered tests.

Phase 1 result (2026-09-28): **passed**. The public tools now expose strict theme read/validate/apply operations; the REST boundary admits only the exact site-scoped `thememanager/ApplyTheme` path and rejects tenant-theme administration and path variants. Graph page authoring now supports typed title regions, full-width and vertical sections, and section emphasis. The real Equipment Portal accepted the palette and enhanced homepage, readback succeeded, and a second blueprint apply reported the theme as a no-op. Arbitrary CSS remains intentionally unsupported.

### Phase 2 — site-data components

- Implement typed Hero, List, Document Library, News, and Events.
- Resolve every SharePoint object by local name and register backend-returned IDs.
- Add full blueprint diff/apply/readback and duplicate prevention.
- Upgrade the Equipment Portal homepage and visually verify desktop rendering.

Phase 2 result (2026-09-28): **passed**. The public page/blueprint surface now accepts typed Hero, List, Document Library, News, and Events definitions. The trusted backend resolves local list/library/view IDs, stores no caller-supplied component GUIDs, and produces deterministic control IDs plus a definition/live fingerprint. The real published Equipment Portal rendered a three-tile Hero at desktop width and SharePoint's responsive carousel at narrow width, the exact Pending Approval list view, Equipment Documentation library, current-site News, and Events. PnP readback found all six intended controls. A second blueprint apply reported the page, publish, and theme as no-ops and retained the same control IDs.

### Phase 3 — external service components

- Implement Microsoft Forms and Power Apps typed properties.
- Implement allowlisted Embed without changing HTML Field Security.
- Add hostile URL/ID tests and live rendering tests using non-sensitive disposable resources that already exist or are supplied by the human.

Phase 3 implementation result (2026-09-28): **partial/experimental**. Microsoft Forms and Power Apps have typed identifier-only serializers, and the local site reports both built-in components available. Embed accepts only an HTTPS URL whose exact fully qualified hostname is in human-managed trusted configuration; localhost, IP literals, credentials, custom ports, fragments, subdomain widening, raw HTML, and scripts are rejected. Unit/adversarial tests pass. Rendered acceptance is not claimed because no disposable Form/App identifiers were supplied and the live configuration intentionally enables no external Embed host.

### Phase 4 — optional SPFx component

Proceed only if native theming and components cannot meet an accepted visual requirement and the human explicitly approves app-catalog deployment/custom-code administration.

## 10. Security acceptance criteria

In addition to the baseline adversarial suite:

- PnP context creation rejects a URL other than the immutable configured site;
- PnP authentication rejects unknown resource audiences and hosts;
- live PnP calls emit no token or raw response containing a token to MCP/audit output;
- target list/library/view/page/file names cannot carry a site path or foreign GUID;
- News and Events cannot select another site, hub, or tenant-wide source;
- Hero links and Embed URLs pass canonical host allowlist validation after repeated decoding;
- redirects are not followed outside the configured SharePoint origin by the trusted backend;
- Power Apps/Form IDs can appear only in component property data, never request paths;
- exact theme endpoint tests prove tenant theme add/update/delete cannot be represented;
- no tool accepts component GUID, raw property JSON, CanvasContent1, CSS, JavaScript, HTML iframe, REST URL, PnP command, or SPFx package;
- configuration and external-host allowlists remain trusted bootstrap state and cannot be changed by MCP.

## 11. Functional acceptance

The upgraded Equipment Request Portal must visibly contain:

- branded site palette;
- accessible title region and full-width Hero with two to five working tiles;
- Pending Approval List web part bound to the exact local view;
- Equipment Documentation library web part bound to the intended library/view;
- current-site News and Events components;
- polished section emphasis, spacing, images, calls to action, and responsive layouts;
- Forms, Power Apps, and Embed acceptance pages when valid external resource identifiers are available.

Verification must include tool readback, second-apply no-op/duplicate checks, browser inspection, and screenshots at desktop plus a narrow responsive viewport. Empty web parts, maintenance-mode errors, placeholder components, or serialized JSON that does not render do not count as success.

## 12. Open decisions after the spike

- Whether PnP Core can save all target web parts with the current `AllSites.Manage` session or a narrower/different SharePoint delegated permission is required.
- Which target web parts have stable enough property contracts for `Complete` versus `Experimental` classification.
- Whether current-theme readback exposes a normalized palette sufficient for exact idempotency; otherwise theme apply will use a locally recorded trusted last-applied hash plus site readback evidence.
- Whether Forms and Power Apps test resources already exist. Creation/administration of those services is not implied by embedding them.
- Whether native SharePoint visuals satisfy the desired finish before considering the separately approved SPFx option.
