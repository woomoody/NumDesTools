---
name: numdestools-feishu-project
description: Query Feishu Project data for NumDesTools with FeishuProjectMcp. Use when you need project info, work item metadata, comments, field config, MQL search results, or download URLs, instead of guessing from memory or searching the codebase.
---

# numdestools-feishu-project

Use this skill when work on NumDesTools depends on Feishu Project data.

## Preferred source

Use the `FeishuProjectMcp` tools first for:

- project info
- work item search
- work item field config
- comments and operation records
- MQL-based lookup
- user or team info

## Local secret rule

- Do not store the Feishu MCP token in repo files
- The real `FeishuProjectMcp` auth lives in local user config
- Repo files may only contain safe templates or dependency metadata

## If the MCP is unavailable

- Check local user-level Codex config for `FeishuProjectMcp`
- For Claude Code, check the local ignored `.mcp.json`
- If the MCP is still unavailable, stop and tell the user what local config is missing
