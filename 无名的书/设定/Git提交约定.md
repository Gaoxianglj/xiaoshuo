# Git 提交约定

> 建立：2026-08-13（作者确认）
> 用途：项目内所有文件改动的提交规则。新 Chat 在任务收尾时必须执行，无需再向作者询问是否提交。

## 1. 自动提交

- 作者要求：**每一轮任务的最后，自动把本轮所有文件改动提交到 git**。写作、改稿、设定整理、追踪同步等任何任务都适用。
- 即使作者没有在消息里提到 git，任务结束时也要按本文件提交，不得跳过。

## 2. 分支规则

- 默认提交分支：`deepseek_harness_MAC`（当前工作分支）。
- **严禁提交到 `main` 和 `deepseek_harness_PC` 分支**。`main` 与 `deepseek_harness_PC` 只由作者本人处理合并与推送。
- 提交前先 `git branch` 确认所在分支是 `deepseek_harness_MAC`，且不是 `main` 或 `deepseek_harness_PC`；若误在其他分支，先切回 `deepseek_harness_MAC` 再提交，不得在 `main` 或 `deepseek_harness_PC` 上产生提交。

## 3. 提交信息

- 使用中文，简明概括本轮改动，例如：“按细纲重新生成第001章正文”“同步第002章快照与上下文”。
- 多条不同性质改动时，可在正文中分点列出主要文件或事项。

## 4. 提交范围

- 只提交本轮任务实际产生的项目文件改动，遵循 `.gitignore`（忽略 `novel-project/`、Office 临时文件等）。
- 不把与任务无关的历史改动混入提交；需要时可用 `git add` 指定路径。

## 5. 推送

- 作者于 2026-08-13 确认：**每次提交后自动推送到远端**，目标为 `origin/deepseek_harness_MAC`。
- 每轮任务收尾时：本地提交 → `git push origin deepseek_harness_MAC`。若推送失败（网络、凭据、冲突等），在回复中明确报告失败原因，不得静默跳过。
- `main` 与 `deepseek_harness_PC` 仍由作者本人处理，不提交、不推送。
