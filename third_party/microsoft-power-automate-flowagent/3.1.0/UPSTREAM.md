# Microsoft Power Automate FlowAgent bundle

- Upstream repository: `https://github.com/microsoft/power-platform-skills`
- Upstream commit: `943401d773b240312c1c05fe83b9e54bbaf9de7f`
- Plugin version: `3.1.0`
- Vendored file: `plugins/power-automate/server/mcp.mjs`
- SHA-256: `85b69e290528c3d3f0d410eaf91c2dcddec64809277a5c54076797d35cc18c47`
- License: MIT; see `LICENSE` in this directory.

The bundle is never registered as a client-facing MCP server. The trusted .NET
process starts it privately and admits only a reviewed internal tool allowlist.
Environment IDs, flow IDs, definitions, and run IDs are resolved and validated
by the .NET capability boundary before the sidecar receives a request.

To update the bundle, review the upstream diff, replace `mcp.mjs`, update the
commit/version/hash above and in `FlowAgentRuntime`, run the boundary/unit suite,
then perform a fresh packaged-client and live Power Automate acceptance test.
