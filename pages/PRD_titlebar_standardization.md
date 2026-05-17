# PRD：青川县园区平台7个子页面标题栏统一

## 1. 项目信息

| 项 | 值 |
|---|-----|
| **语言** | zh-CN |
| **技术栈** | 纯 HTML + CSS + JS（无框架） |
| **项目名称** | park_titlebar_standardization |
| **原始需求** | 将 park_*.html 7个页面的标题栏、按钮、底部切换卡片样式完全统一，消除不一致 |

---

## 2. 产品定义

### 2.1 Product Goals

1. **零视觉差异**：7个子页面在任意浏览器/分辨率下，标题栏、按钮、底部切换卡片的视觉效果完全一致
2. **单一事实源（Single Source of Truth）**：共享 CSS（park_common.css）是标题栏/按钮/底部卡片样式的唯一定义，各页面不重复声明
3. **修改一次、全局生效**：调整任意标题栏/按钮/底部卡片样式只需修改 park_common.css，无需逐页修改

### 2.2 User Stories

1. As a **平台运维人员**, I want 修改一次 CSS 就能让所有7个页面标题栏同时更新 so that 维护效率高、不会遗漏
2. As a **园区管理者**, I want 在任意子页面之间切换时标题栏和按钮位置/样式保持一致 so that 使用体验流畅、平台显得专业
3. As a **前端开发者**, I want 清晰的 CSS 架构和唯一的类名约定 so that 新增页面时不会引入不一致
4. As a **QA 测试人员**, I want 明确的验收标准 so that 可以通过简单的视觉检查确认一致性

---

## 3. 技术规范

### 3.1 Requirements Pool

#### P0（Must have — 阻塞性问题）

| ID | 需求 | 说明 |
|----|------|------|
| P0-1 | **park_energy.html 标题栏修复** | 将 `class="title-bar"` 改为 `class="title-box"`，按钮容器使用 `class="title-right"`，时间用 `class="title-left"` |
| P0-2 | **park_common.css 补充缺失 CSS 变量** | 添加 `--text-primary`, `--sidebar-width`, `--split-color`, `--success-color`, `--text-tertiary` |
| P0-3 | **park_common.css 补充 .time 独立选择器** | 当前 `.time` 样式写在 `.title-box .time` 下，需要增加独立 `.time` 选择器供非 title-box 场景 |
| P0-4 | **park_safety.html 验证修复** | 确保 park_safety.html 引用的所有 CSS 变量在 park_common.css 中已定义 |

#### P1（Should have — 消除冗余）

| ID | 需求 | 说明 |
|----|------|------|
| P1-1 | **删除 park_sandbox.html 中重复 CSS** | 删除 `.map-wrapper`, `#amap`, `.control-bar`, `.ctrl-btn` 等已在 park_common.css 中定义的样式（约 50 行） |
| P1-2 | **删除 park_land.html 中重复 CSS** | 删除 `.control-bar`, `.ctrl-btn` 等已在 park_common.css 中定义的样式 |
| P1-3 | **各页面删除重复的 bottom-tabs CSS** | park_economy/enterprise/service/energy/land/safety/sandbox 全部7个页面均重复定义了 `.bottom-tabs` `.tab-btn` 及动画，需删除 |

#### P2（Nice to have — 架构优化）

| ID | 需求 | 说明 |
|----|------|------|
| P2-1 | **各页面统一 time ID** | sandbox 用 `currentTime`，其他用 `headerTime`，energy 用 `energyTime` → 统一为 `currentTime` |
| P2-2 | **统一 loading-overlay 位置** | economy/enterprise/service 的 loading-overlay 放在了 title-box 之后（DOM 结构异常），应移到 screen-container 之前 |

---

### 3.2 UI Design Draft

#### 3.2.1 标题栏标准结构（所有页面必须遵循）

```html
<div class="screen-container">
    <!-- 顶部标题栏 -->
    <div class="title-box">
        <div class="title-left">
            <span class="time" id="currentTime"></span>
        </div>
        <div class="title-center">
            <h1>页面标题</h1>
        </div>
        <div class="title-right">
            <button class="common-btn" onclick="refreshData()" title="刷新数据">
                <i class="fa fa-refresh"></i>
            </button>
            <!-- 可选：页面专属按钮 -->
            <a href="../index.html" class="back-btn" title="返回主页">
                <i class="fa fa-home"></i>
            </a>
        </div>
    </div>
    <!-- 页面主体内容 -->
</div>
```

#### 3.2.2 按钮类名规范

| 类名 | 用途 | 规格 |
|------|------|------|
| `common-btn` | 功能操作按钮（刷新、宫格切换等） | 38×38px, font 16px, 圆角8px |
| `back-btn` | 导航跳转链接（返回主页、跳转子页） | 38×38px, font 16px, 圆角8px |
| `ctrl-btn` | 地图/视频控制按钮（图层、2D/3D等） | 44×44px, font 18px, 圆角8px |
| `tab-btn` | 底部圆形切换标签 | 80×80px, font 27px, 圆形 |

#### 3.2.3 底部切换卡片标准结构

```html
<div class="bottom-tabs">
    <a class="tab-btn" href="park_sandbox.html">沙盘</a>
    <a class="tab-btn" href="park_economy.html">经济</a>
    <a class="tab-btn" href="park_land.html">用地</a>
    <a class="tab-btn" href="park_enterprise.html">企业</a>
    <a class="tab-btn" href="park_service.html">服务</a>
    <a class="tab-btn" href="park_energy.html">能耗</a>
    <a class="tab-btn active" style="pointer-events:none;opacity:0.8">安全</a>
</div>
```

> **规则**：当前页面标签加 `class="tab-btn active"` 和 `style="pointer-events:none;opacity:0.8"`，其他标签仅 `class="tab-btn"`。

---

### 3.3 park_common.css 需要新增的内容

```css
/* 新增 CSS 变量 */
:root {
    --text-primary: #e0ecff;
    --text-tertiary: #557199;
    --sidebar-width: 260px;
    --split-color: rgba(43,104,224,0.35);
    --success-color: #4cd964;
}

/* 新增独立 .time 样式 */
.time {
    font-size: 14px;
    color: var(--text-secondary);
}
```

---

### 3.4 各页面改造清单

| 页面 | 修复动作 |
|------|---------|
| **park_energy.html** | ① `title-bar` → `title-box`；② 按钮容器用 `title-right`；③ 时间用 `title-left`；④ `eng-map` 的 `top:64px` 调整为使用 grid 布局 |
| **park_sandbox.html** | 删除重复的 `.map-wrapper`, `#amap`, `.control-bar`, `.ctrl-btn`, `@keyframes blinkMarker`, `@keyframes blinkBorder` |
| **park_land.html** | 删除重复的 `.control-bar`, `.ctrl-btn` |
| **park_economy.html** | 删除重复的 `.bottom-tabs` / `.tab-btn` 全部 CSS（约75行） |
| **park_enterprise.html** | 同上 |
| **park_service.html** | 同上 |
| **park_safety.html** | 删除重复的 `.bottom-tabs` / `.tab-btn` 全部 CSS + 删除 `.time` 样式 |

---

### 3.5 验收标准

| # | 标准 | 验证方法 |
|---|------|---------|
| AC-1 | 7个页面标题栏高度、背景渐变、边框、阴影完全一致 | 分别在 Chrome/Edge 打开7个页面，截图叠加对比 |
| AC-2 | 所有 `common-btn` 按钮 38×38px、字体16px、圆角8px | DevTools 检查 computed styles |
| AC-3 | 所有 `back-btn` 链接 38×38px、字体16px、圆角8px | DevTools 检查 computed styles |
| AC-4 | 底部7个圆形标签各页面位置、大小、动画完全一致 | 截图叠加对比 |
| AC-5 | park_common.css 是标题栏/按钮/底部卡片的唯一样式来源 | 搜索所有 park_*.html 中不应再包含 `.title-box{` `.bottom-tabs{` `.tab-btn{` `.common-btn{` `.back-btn{` 的重复定义 |
| AC-6 | park_energy.html 与其他页面视觉一致（之前异常） | 对比 energy 页面与 sandbox/land 页面截图 |
| AC-7 | 所有页面功能不受影响（刷新、跳转、时间显示） | 逐页手动测试 |
| AC-8 | 7个页面底部的标签顺序一致：沙盘→经济→用地→企业→服务→能耗→安全 | 目视检查 |

---

### 3.6 Open Questions

1. **park_energy.html 的标题栏使用绝对定位而非 grid 布局**：如果改为 `title-box` + `screen-container` grid，需要调整 `.eng-map` 的定位方式（它目前 `top:64px` 硬编码）。是否需要一并改为 grid 布局？
2. **time ID 统一**：sandbox 用 `currentTime`，economy/enterprise/service/safety 用 `headerTime`，energy 用 `energyTime`。是否需要全部统一为一个 ID（如 `currentTime`）以简化 JS？
3. **loading-overlay DOM 位置异常**：economy/enterprise/service 三页的 loading-overlay 放在了 `</div>` title-box 之后但在 map-wrapper 之前，导致它在 screen-container 内部而非全屏覆盖。是否需要修复？
