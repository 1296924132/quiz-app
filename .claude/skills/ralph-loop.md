---
name: ralph-loop
description: Ralph loop — 自治 AI 开发循环，反复运行直到所有 PRD 事项完成
---

# Ralph Loop (自治开发循环)

Ralph 是一个自治 AI 代理循环，每次以**全新上下文**启动 Claude Code，反复执行直到所有 PRD 需求完成。避免了上下文膨胀导致的幻觉问题。

## 快速开始

### 1. 初始化项目

```bash
ralph init
```

这会创建 `prd/` 目录及必要的配置文件。

### 2. 编写 PRD

在 `prd/prd.md` 中编写产品需求文档。

### 3. 运行循环

```bash
ralph run
```

或指定最大迭代次数：

```bash
ralph run 20
```

### 4. 查看状态

```bash
ralph status      # 查看当前进度
ralph inspect     # 查看运行状态
ralph history     # 查看历史记录
ralph reset       # 重置迭代计数器
```

## 工作流程

1. **编写 PRD** — 在 `prd/prd.md` 中定义需求
2. **初始化** — `ralph init` 生成配置文件
3. **运行** — `ralph run` 启动循环，每次迭代：
   - 读取当前进度
   - 选择下一个未完成的用户故事
   - 启动 Claude Code 以全新上下文实现
   - 运行质量检查
   - 自动提交并标记完成
   - 重复直到全部完成

## 输出

每次迭代的结果会保存在 `ralph/` 目录中，包括进度日志、迭代记录等。
