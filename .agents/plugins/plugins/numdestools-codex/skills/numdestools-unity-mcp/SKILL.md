---
name: numdestools-unity-mcp
description: Orchestrate Unity Editor workflows through the local unity-mcp server. Use when you need to inspect Unity editor state, read the Unity console, launch Unity-side automation flows, edit scenes, manage GameObjects, modify scripts, run tests, or otherwise drive a Unity project through MCP instead of guessing or editing blind.
---

# numdestools-unity-mcp

Use this skill when a task depends on a live Unity Editor connected through `unity-mcp`.

## Core workflow

1. Check the MCP connection and editor state first.
2. Read Unity state/resources before mutating anything.
3. Prefer discovery before action:
   - inspect editor state
   - inspect scene / objects
   - inspect console
4. Only then apply scene, script, or package changes.
5. Verify with console output and, if available, Unity-side screenshots.

## Good defaults

- Treat the Unity editor as the source of truth for scene state.
- After script changes, wait for compilation to finish before doing more work.
- After scene changes, verify via console and screenshots instead of assuming success.
- For repeated or multi-object work, prefer batched operations.

## Typical tasks

- Read console errors and warnings
- Inspect whether the editor is compiling or blocked
- Modify GameObjects or components
- Create or update scripts, then verify compilation
- Run Unity tests
- Drive visual validation with screenshots

## Local dependency rule

- This skill depends on the local `unity-mcp` server configured in user-level Codex config.
- Do not store Unity MCP auth or localhost runtime wiring in repo files beyond safe dependency metadata.
- If the MCP is unavailable, stop and report that the local Unity MCP server is not connected.
