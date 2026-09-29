# MCP tool reference

The server registers 132 typed tools across three service classes. SharePoint is always the immutable configured site. Power Automate, OneDrive, and mail exist only when a human has saved an explicit resource policy and enabled the matching capability. No tool accepts a tenant, site ID, arbitrary absolute URL, Graph/REST path, HTTP method, token, command, or script.

| Family | Tools | Boundary and behavior |
|---|---|---|
| Identity/scope | `who_am_i`, `get_configured_scope` | Delegated identity, site-effective permission bits, immutable aliases/scopes, and capability flags; secrets are omitted. |
| Site | `site_get_info`, `site_get_capabilities`, `site_inventory`, `site_validate_health` | Read only the configured site and report honest support states. |
| Appearance | `site_appearance_get`, `site_appearance_update` | Title, description, logo, header/footer, navigation style, audience targeting, welcome page, ETag; writes require `siteSettingsWrite` and default to dry-run. |
| Themes | `site_theme_get`, `site_theme_list`, `site_theme_validate`, `site_theme_apply` | Site-visible themes and strict semantic palettes. Only this site's apply endpoint is reachable; tenant theme administration is not. |
| Lists | `list_list`, `list_get`, `list_create`, `list_update`, `list_delete` | Graph lifecycle plus site settings; writes require `listsWrite`; delete defaults to dry-run. |
| Columns | `column_list`, `column_get`, `column_create`, `column_update`, `column_delete` | Typed definitions only; no field XML. Incompatible type changes are rejected. |
| Items | `item_get`, `item_query`, `item_create`, `item_update`, `item_delete`, `item_recycle`, `item_version_list`, `item_version_restore` | Typed filters, paging, selected/expanded fields, partial patches, ETag concurrency, recycle, and recoverable versions. |
| Views | `view_list`, `view_get`, `view_create`, `view_update`, `view_delete`, `view_set_default` | Selected columns, typed filter/sort/grouping, row limit, formatting, and default selection. The server generates CAML; callers cannot supply it. |
| Forms | `form_get`, `configure_form`, `field_visibility_get`, `field_visibility_set`, `form_format_set` | Header/body/footer, sections, field order/read-only state, and separate New/Edit/Display/conditional visibility. |
| Formatting | `column_format_set`, `column_format_clear`, `view_format_set`, `view_format_clear` | Official declarative SharePoint JSON after size and active-content validation. |
| Libraries | `library_list`, `library_get`, `library_create`, `library_update`, `library_delete` | Site-local libraries; writes require `documentsWrite`; delete defaults to dry-run. |
| Files/folders | `file_list`, `file_get_metadata`, `file_read`, `file_write`, `file_upload`, `file_delete`, `file_move`, `file_copy`, `file_copy_if_absent`, `file_recycle`, `file_version_list`, `file_version_restore`, `folder_create`, `folder_list`, `folder_delete`, `folder_move` | Exact site-library paths, resumable large uploads, create-only copy, versions, and recycle. Download/upload URLs never leave the process. |
| Evidence | `evidence_search`, `evidence_read_text` | Searches names and locally extracts text from text/OpenXML files in the configured site's libraries only. |
| Content types | `content_type_list`, `content_type_create`, `content_type_update`, `content_type_add_to_list` | Site/list-local only; no tenant content-type hub administration. |
| Pages | `page_list`, `page_get`, `page_create`, `page_update`, `page_delete`, `page_publish`, `page_unpublish`, `page_set_home`, `page_copy_if_absent`, `page_update_draft`, `page_version_list`, `page_version_restore`, `page_recycle` | Typed canvas, draft/published lifecycle, clone-if-absent, ETag draft replacement, versions, recycle, and welcome page. Each write has a separate capability gate where appropriate. |
| Sections | `page_section_add`, `page_section_update`, `page_section_delete` | Reconciles the complete typed canvas; destructive removal supports dry-run. |
| Web parts | `webpart_available`, `webpart_get`, `webpart_add`, `webpart_update`, `webpart_move`, `webpart_delete` | Friendly component models. Graph-supported parts have standalone operations; PnP-only advanced parts are reconciled through a complete page/blueprint. Raw component IDs/properties are not accepted. |
| Navigation | `navigation_get`, `navigation_add`, `navigation_update`, `navigation_remove`, `navigation_reorder` | Quick Launch, top navigation, footer, nesting, ordering, new-window state where SharePoint exposes it, and audience IDs. Targets remain site-local. |
| Permissions | `permissions_inspect`, `permissions_grant`, `permissions_revoke`, `permissions_reset_inheritance` | Site/list/library/page-local target shapes only. Mutation additionally requires both `permissionsMutationEnabled` and `permissionsWrite`; backend support remains partial. |
| Blueprints/change control | `site_validate_blueprint`, `site_plan_blueprint`, `site_diff_blueprint`, `site_apply_blueprint`, `preview_change`, `snapshot_resource` | Validate, inspect, diff, dry-run, apply, read back, verify, and snapshot versions/state. Stable identities prevent duplicate resources. |
| Power Automate | `flow_get_status`, `flow_list`, `flow_get`, `flow_export`, `flow_update`, `flow_clone`, `flow_import`, `flow_set_state` | Exact configured environment and flow aliases only. Saved definitions, connection references, expected versions, validation, state, and readback. Generic HTTP/custom connectors and foreign site/resource IDs are rejected. |
| Flow runs | `flow_run_list`, `flow_run_get`, `flow_run_actions_get`, `flow_run_resubmit`, `flow_run_cancel` | Exact configured flow only; opaque run IDs; action/trigger inputs, outputs, errors, and nested details. Signed content URLs are consumed internally and stripped. |
| OneDrive evidence | `onedrive_get_status`, `onedrive_search`, `onedrive_read_text` | Optional fixed drive and root item; read-only; no caller-selectable drive or root. |
| Mail | `mail_get_status`, `mail_search`, `mail_get`, `mail_attachment_list`, `mail_attachment_read`, `mail_draft_create`, `mail_draft_send` | Optional fixed folder aliases. Read, draft, and send are separate flags; send accepts only runtime-created drafts and trusted recipient domains. |

## Safe item queries

`item_query` accepts typed predicates (`equals`, `notEquals`, comparisons, `contains`, `startsWith`, `isNull`, `isNotNull`), bounded page sizes, continuation tokens created by the server, selected fields, lookup expansion, and typed sorting. It never accepts OData or CAML text. Updates patch only supplied fields and may require the current ETag.

## Path and identifier rules

SharePoint files use `Library Title/folder/file.ext`. Absolute/rooted paths, URLs, hostnames, `sites/`, backslashes after decoding, dot segments, encoded traversal, and foreign GUIDs are rejected. External resources use trusted aliases; IDs returned by a bound backend are registered before reuse. Power Automate pagination must remain under the exact configured environment/flow route.

## Capability policy

New authority domains fail closed. A human uses trusted `configure-scope --policy config/runtime-policy.example.json` to select resource aliases and enable writes. Normal MCP calls cannot run that command or modify configuration. Destructive operations default to `dryRun=true` where supported, and successful mutations are appended to the trusted JSONL audit log.

## Current limitations

Power Automate uses a pinned private Microsoft FlowAgent sidecar authenticated by the user's Azure CLI session. Environment and connector discovery remain bootstrap-only; runtime tools accept configured aliases, never environment IDs, arbitrary flow IDs, or caller-selected connection IDs. Live creation/readback, state change, SharePoint polling trigger, item update, run listing, and action inspection passed against the configured sandbox. Resubmit/cancel still lack an appropriate failed/running live fixture. Taxonomy, list-item attachments, reliable permission mutation, and standalone edits to PnP-only web parts remain partial or blocked. Unsupported operations return an explicit error; they never report synthetic success.
