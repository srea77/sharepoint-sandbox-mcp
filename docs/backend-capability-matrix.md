# Backend capability matrix

This is the implemented state, not a theoretical API inventory. `Complete` means the implementation and live sandbox path passed. `Experimental` means code and automated tests exist but the exact service/resource path lacks live acceptance. `Blocked` returns an explicit error rather than claiming success.

| Capability | Status | Backend | Evidence / limitation |
|---|---|---|---|
| Identity, effective site permissions, immutable scope | Complete | SharePoint REST + core policy | Live site identity and permission bits; no credentials or site selector. |
| Site inventory/health | Complete | Graph + REST + PnP Core | Live lists, libraries, views, forms, pages, navigation, features, and component availability. |
| Site settings/appearance | Partial/experimental | SharePoint REST | Typed title/description/logo/header/footer/nav-style/audience/welcome-page models, ETags, dry-run, audit. Some tenant header/footer combinations need live coverage. |
| Site theme list/apply | Complete/experimental | SharePoint REST + trusted state | Exact site theme endpoints, strict palette, live apply/readback/no-op. No tenant theme create/update/delete. |
| List/library lifecycle | Complete | Graph + REST | CRUD, settings, versioning, and attachments setting; live acceptance. |
| Columns | Complete except taxonomy | Graph | Text/multiline/number/currency/date/boolean/choice/multichoice/person/lookup/hyperlink/calculated; get/update/delete and no-op comparison. Taxonomy remains blocked. |
| Records | Complete | Graph | Typed filters, paging, selected fields, lookup expansion, partial patch, ETags, delete/recycle, versions/restore. Live CRUD coverage remains narrower than structural portal acceptance. |
| Item attachments | Not implemented | — | Document APIs do not stand in for list-item attachments. |
| Views | Complete | SharePoint REST | Typed filters/sorts/groups, fields, row limits, default, formatting, readback. |
| Form layout/visibility | Complete/experimental | SharePoint REST | Header/body/footer, sections/order, read-only fields, separate New/Edit/Display and conditional rules. Layout sections rendered live; conditional-rule variants are automated-test verified. |
| Formatting | Complete | SharePoint REST | Column/view/form declarative JSON with validation; portal status formatting verified live. |
| Files/folders | Complete | Graph | Exact paths, upload/download, list, move/copy, create-if-absent, recycle, versions. Large files use resumable upload sessions with bounded Microsoft upload origins. |
| Site evidence extraction | Complete/experimental | Graph + local OpenXML | Fixed-site library search and local extraction for text/DOCX/PPTX/XLSX; no tenant search. |
| Content types | Partial/experimental | Graph | Site/list enumerate/create/update/add-copy; no tenant hub. |
| Modern page lifecycle | Complete/experimental | Graph + PnP Core + REST | Typed create/update/draft/copy, publish/unpublish, home-page, versions/restore, recycle. Create/publish/home/readback/idempotency live verified; unpublish/restore remain policy-gated variants without destructive live acceptance. |
| Sections/title region | Complete/experimental | Graph + PnP Core | Common/full-width/vertical layouts, emphasis and title regions; live rendered acceptance. |
| Graph web parts | Partial/experimental | Graph | Text plus supported standard parts and standalone lifecycle where Microsoft Graph supports them. |
| Hero/List/Library/News/Events | Complete/experimental | PnP Core + bound resolution | Typed, deterministic, site-local, rendered live and second-apply no-op verified. |
| Forms/Power Apps/Embed components | Partial/experimental | PnP Core | Typed and boundary-tested. No disposable live form/app/embed target was available. |
| Standalone PnP-only part mutation | Blocked/partial | PnP Core | Safely reconciled only through the complete page/blueprint contract. |
| Navigation | Complete/experimental | SharePoint REST | Quick Launch/top/footer, nesting/order/audiences and MenuState new-window handling; Quick Launch live read/write/idempotency verified. Footer new-window mutation is explicitly unsupported. |
| Permission inspection | Partial/experimental | Graph/REST | Typed site-local resources only; depends on delegated visibility. |
| Permission mutation | Blocked/partial | — | Double-gated but no backend is claimed complete until inheritance semantics receive live adversarial acceptance. |
| Blueprint/diff/idempotency | Complete for implemented families | Core + composite backend | Equipment Portal dry-run/apply/readback/second apply passed live without duplicate named resources or controls. |
| Snapshot/preview/audit | Complete | Core + bound services + JSONL | Typed snapshots/versions, dry-run plans, resulting IDs/ETags where returned, fresh readback, redacted mutation log. |
| Power Automate definitions/state | Complete/experimental | Pinned Microsoft FlowAgent behind bounded adapter | Azure CLI browser authentication, environment discovery, stopped flow creation, policy aliasing, and public-MCP readback passed live. Definitions/import/clone/update/state use connector/resource validation, preflight/preview, expected-version checks, audit, and fresh readback. Updates/state changes were unit-tested but not destructively exercised on an organizational production flow. |
| Power Automate runs/actions | Complete/experimental | Pinned Microsoft FlowAgent behind bounded adapter | Fixed-flow run list/details/actions and state changes are implemented and adversarially tested. Live acceptance enabled a configured SharePoint-triggered flow, observed a successful `PatchItem`, read its run/action details through MCP, and stopped it again. Resubmit/cancel remain implemented but lack a suitable failed/running live fixture. FlowAgent currently caps history at 50 and exposes no safe cursor input. |
| Fixed-root OneDrive evidence | Experimental; not configured live | Microsoft Graph | Human-selected drive/root; read-only search and local text extraction. Disabled by default. |
| Fixed-folder mailbox evidence | Experimental; not configured live | Microsoft Graph | Human-selected folder aliases; read/search/attachments. Disabled by default. |
| Mail draft/send | Experimental; not configured live | Microsoft Graph | Separate default-off flags, trusted recipient domains, runtime-created drafts only, dry-run. |

The SharePoint Graph transport is permanently rooted at `/v1.0/sites/{configured-id}/`; SharePoint REST is permanently rooted at `{configured-site}/_api/`; PnP Core accepts only the same SharePoint site and fixed Graph site resource. Power Automate's private adapter always supplies one trusted environment and resolves only explicit flow aliases or IDs created by that same runtime; raw FlowAgent discovery and generic operations are absent from the runtime allowlist. OneDrive/mail transports likewise use fixed trusted identifiers. All first-party handlers disable redirects or validate the exact continuation/upload origin. PnP PowerShell is inactive.
