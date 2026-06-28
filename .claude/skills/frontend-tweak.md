---
name: frontend-tweak
description: 前端调试优化 — 快速修改刷题神器的前端样式、交互、功能
---

# 前端调试优化大师

## 适用场景

- 修改 `刷题神器/公文写作刷题神器（离线版）.html` 的样式、交互逻辑、功能
- 添加新功能（如新题型、统计面板、快捷键）
- 修复 UI bug 或优化用户体验
- 保持与现有代码风格一致

## 核心原则

1. **先读代码，再动手** — 必须先理解现有结构和风格，不做风格割裂的修改
2. **就地修改** — 刷题神器是单个 HTML 文件，所有 CSS/JS 都在一个文件内
3. **增量修改** — 每次只改一个功能点，改完立刻验证
4. **保持兼容** — 不破坏 localStorage 数据格式（用户刷题进度存在本地）

---

## 项目结构速览

```
刷题神器/
├── 公文写作刷题神器（离线版）.html  # 主应用（单文件，CSS+JS 内联）
├── 刷题神器.html                    # 轻量加载版（需同目录 JSON 题库）
└── index.html                       # GitHub Pages 入口
```

## 现有代码风格约定

### CSS
- 使用 CSS 变量统一管理主题色（`:root` 中定义）
- 深色主题背景，光晕/发光效果
- 卡片式设计，hover 有动效
- 已用 `@media` 做响应式适配

### JavaScript
- 原生 JS（无框架）
- 数据存在 `localStorage` 中
- 事件用 `addEventListener` 绑定
- DOM 操作用原生 API

---

## 常见修改模式

### 1. 修改样式（CSS）

```css
/* 找到对应的 CSS 变量或类名 */
:root {
    --primary-color: #ff6b35;   /* 修改主题色 */
}

/* 或直接修改选择器 */
.quiz-card {
    border-radius: 12px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
}
```

### 2. 添加交互（JavaScript）

```javascript
// 先检查功能是否存在
if (typeof existingFunction === 'undefined') {
    // 新功能：用已有的命名风格
    function handleNewFeature() {
        // ...
    }
    document.addEventListener('DOMContentLoaded', handleNewFeature);
}
```

### 3. 新增 DOM 元素

```html
<!-- 用与现有元素一致的 class 命名和 HTML 结构 -->
<div class="feature-panel">
    <h3 class="panel-title">新功能标题</h3>
    <div class="panel-content">
        <!-- content -->
    </div>
</div>
```

---

## 调试流程

```bash
# 1. 打开文件在浏览器预览（验证修改效果）
start 刷题神器/公文写作刷题神器（离线版）.html

# 2. 检查控制台有无报错
# 3. 检查 localStorage 数据完整性
# 4. 确认响应式布局正常
```

## 安全清单

每次修改后检查：
- [ ] 控制台无报错
- [ ] `localStorage` 数据读写正常
- [ ] 现有功能不受影响（答题、切换模式、查看解析）
- [ ] 页面布局没有塌陷/溢出
- [ ] 暗色主题一致，没有突兀的亮色块

---

## 工作流

```
用户需求（"把按钮改成蓝色" / "加个计时器"）
  ↓
[1] Read 目标 HTML 文件（先理解结构和风格）
  ↓
[2] 定位要改的代码段
  ↓
[3] 改完 → 浏览器刷新验证
  ↓
[4] 跑安全检查清单
  ↓
[5] 完成
```
