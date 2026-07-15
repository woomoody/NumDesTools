# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

通用工程规则（构建命令、架构、数据约定、编码规范、Excel读写规范、输出目录规范）见 [AGENTS.md](AGENTS.md)——
那份是给任何在这个仓库里写代码的 agent（包括 Codex）看的，不分工具。本文件只记 Claude Code 专属行为。

### 输出目录规范 · CC 专属操作习惯

`OutputRootPath`（默认 `Documents\NumDesOutput\`，本地独立 git 仓库不推送）下写文件的子目录规范
见 AGENTS.md。CC（我）手动写文件时，直接写到对应子目录，**写完立即执行**：

```bash
git -C "C:\Users\cent\Documents\NumDesOutput" add -A
git -C "C:\Users\cent\Documents\NumDesOutput" diff --cached --quiet || git -C "C:\Users\cent\Documents\NumDesOutput" commit -m "[描述] 说明内容"
```

### 与 Codex 协作

- 实时交接协议见 `docs/agent-handoff.md`
- 运行态交接文件统一放 `.remember/agent-handoff/<task-name>/`，不提交
- 值得长期保留的结论必须回写到版本化文件，不要只留在交接文件里
- 没有具体任务前不要预建 worktree；并行改动时再按 `docs/agent-handoff.md` 建

## 多模型自动路由

路由总表与 Workflow 规则见全局 `~/.claude/CLAUDE.md`，项目级不重复。此处仅记项目相关要点：

- LiteLLM 网关：`https://litellm.solotopia.net/v1/chat/completions`，Key 见 `ANTHROPIC_AUTH_TOKEN` 环境变量。
- **中文游戏配置分析**（数值合理性、配置审查）用当前会话模型（sonnet/opus），有项目 Memory 和设计规范加成，实测优于 qwen。
- 需要读写文件/代码/git 的任务留在当前会话，其他模型无 CC 工具权限。

<!-- OPENWIKI:START -->

## OpenWiki

This repository uses OpenWiki for recurring code documentation. Start with `openwiki/quickstart.md`, then follow its links to architecture, workflows, domain concepts, operations, integrations, testing guidance, and source maps.

The scheduled OpenWiki GitHub Actions workflow refreshes the repository wiki. Do not hand-edit generated OpenWiki pages unless explicitly asked; prefer updating source code/docs and letting OpenWiki regenerate.

<!-- OPENWIKI:END -->
