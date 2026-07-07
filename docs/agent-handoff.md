# Agent 协作交接协议

这是协议 + 模板，不只是说明文档。真正开始 Claude Code + Codex 协作任务时，按下面的目录和文件职责执行。
**没有具体任务前不要预先建分支/worktree**，建了也是空对象。

## 谁干什么

- Claude Code：长上下文规划、任务拆分、写文档/接口草案/测试思路
- Codex：直接进仓库读代码、改代码、跑命令、修 bug、落地实现
- 两者是独立进程，不共享上下文；需要靠文件交接，而不是假设对方“已经知道”

## 通信文件链路

实时交接文件统一放到：

```text
.remember/agent-handoff/<task-name>/
```

推荐文件结构：

```text
.remember/agent-handoff/<task-name>/
  task.md
  to-codex.md
  to-claude.md
  status.md
  worktree.md
```

说明：

- `.remember/` 已被 git 忽略，运行态交接文件不提交
- 需要长期保留的结论，必须回写到版本化文件（代码、`docs/`、`AGENTS.md`、测试）
- `worktree.md` 只有在当前任务真的用了 worktree 时才创建；否则不要预建

## 每个文件写什么

### `task.md`

稳定上下文的唯一入口。任务开始时建，后续只增量更新。

应该包含：

- 目标
- 当前 owner（`claude` / `codex` / `human`）
- 当前分支或 `TBD`
- 文件归属
- 共享高风险区
- 验收条件

### `to-codex.md`

发给 Codex 的最新动作请求。由 Claude 或人工覆盖写入，不做长历史堆积。

应该包含：

- 当前要 Codex 做什么
- 相关文件/提交/hash
- 不该碰的区域
- 预期验证方式

### `to-claude.md`

发给 Claude 的最新动作请求。由 Codex 或人工覆盖写入，不做长历史堆积。

应该包含：

- 当前要 Claude 做什么
- 需要 Claude 补的背景、设计、文档或 review
- 关键实现结论和风险

### `status.md`

追加式状态日志，按时间倒序写。每次交接前后都更新这里。

每条至少写：

- 时间
- 谁做了什么
- 验证结果
- 下一步谁接手

### `worktree.md`

只在任务需要并行改动时创建。记录：

- `claude/<task-name>` 分支与路径
- `codex/<task-name>` 分支与路径
- 如有需要，再加 `integrate/<task-name>`
- 冲突收口负责人

## 交接顺序

1. 具体任务确定后，新建 `.remember/agent-handoff/<task-name>/`
2. 先写 `task.md`，把目标、边界、文件归属写清楚
3. 发起方先更新 `status.md`，再覆盖写目标方的 `to-*.md`
4. 接手方完成动作后，先更新 `status.md`，如需回交再覆盖对方的 `to-*.md`
5. 需要长期保留的信息，不留在 `.remember/` 里，必须回写到版本化文件

## Worktree 规则

真正要并行改代码时才建：

```bash
cd C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools
git worktree add ..\NumDesTools-claude -b claude/<task-name> master
git worktree add ..\NumDesTools-codex  -b codex/<task-name>  master
```

规则：

- 没有具体任务，不建 worktree
- 只有两边都要改代码，或需要隔离冲突时，才建
- 只有两边都写了大量代码且改动区重叠时，才开 `integrate/<task-name>`

## 合并规则

- Claude 的文档/接口提交 → `git cherry-pick <sha>` 进 Codex 分支
- 最终只从 Codex 分支合回 `master`
- 不要在 `claude/*` 和 `codex/*` 分支里互相 merge，避免冲突扩散
- 冲突只在一个地方解决：谁负责收口，就只在那个 worktree 里改冲突文件

## 模板

### `task.md`

```md
# Task: <task-name>

Status: active
Current owner: <claude|codex|human>
Current branch: <branch-name|TBD>

## Goal
- <goal>

## File ownership
- Claude: <paths>
- Codex: <paths>
- Shared risky paths: <paths>

## Acceptance
- <check 1>
- <check 2>
```

### `to-codex.md` / `to-claude.md`

```md
# Handoff

Updated: <yyyy-mm-dd hh:mm>
Requester: <claude|codex|human>

## Need
- <action item>

## Inputs
- <file/path/commit>

## Do not touch
- <path>

## Verify
- <expected verification>
```

### `status.md`

```md
# Status

- <yyyy-mm-dd hh:mm> <who>: 做了什么；验证结果；下一步谁接手
- <yyyy-mm-dd hh:mm> <who>: 做了什么；验证结果；下一步谁接手
```
