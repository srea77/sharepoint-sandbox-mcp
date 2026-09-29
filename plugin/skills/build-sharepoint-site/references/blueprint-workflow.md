# Blueprint workflow

Use the blueprint tools in this order:

1. `site_plan_blueprint` to inspect current state and normalize intent.
2. `site_validate_blueprint` to reject invalid names, paths, formatting, external
   embed hosts, and unsupported definitions before mutation.
3. `site_diff_blueprint` to present creates, updates, deletions, no-ops, risk, and
   unsupported operations.
4. `site_apply_blueprint` with `dryRun: true`.
5. `site_apply_blueprint` with `dryRun: false` after the plan matches the request.
6. `site_inventory` plus targeted list/page/form/navigation reads for verification.
7. Apply or diff the same blueprint again and verify no duplicate logical resources
   are proposed or created.

Use stable keys for pages, sections, web parts, and navigation. Reference lists,
libraries, and views by site-local names. Use relative library/file paths only.
Formatting may contain supported declarative SharePoint JSON, but never active
script content. External embeds must already be allowed by trusted human
configuration.

When partial support prevents one blueprint operation, report that operation and
continue with independent safe operations. Never broaden authority or switch sites
to work around a limitation.
