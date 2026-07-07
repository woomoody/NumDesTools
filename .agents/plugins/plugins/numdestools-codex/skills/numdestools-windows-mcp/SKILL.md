---
name: numdestools-windows-mcp
description: Use the local windows-mcp server for machine-level Windows inspection and automation when a task needs screenshots, desktop state, or OS-side interaction that is outside normal repo editing.
---

# numdestools-windows-mcp

Use this skill when Codex needs Windows desktop context rather than just repository files.

## Use cases

- Take screenshots
- Inspect desktop/UI state
- Wait for UI transitions
- Support workflows that depend on live Windows UI feedback

## Local dependency rule

- This skill depends on the local `windows-mcp` server configured in user-level Codex config.
- Keep local runtime wiring out of repo files except for safe dependency metadata.
- If the MCP is unavailable, report that local Windows MCP is not available.
