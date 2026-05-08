# Task Pane Current Behavior

日期：2026-04-24

状态：已实现并可联调。当前任务窗格固定 UI、前端系统消息与浏览器预览模式都已接入中英文双语切换，但 AI 自由回复仍保持“尽量跟随用户输入语言”的策略。

## 1. 模块范围

Task Pane 指 Excel 右侧由 WebView2 承载的 React 界面，以及浏览器预览模式下的同一套前端。

当前快照覆盖：

- 任务窗格固定 UI 文案
- 前端生成的系统消息与确认卡片
- 宿主通过 `bridge.getHostContext` 暴露的 UI 语言上下文
- 浏览器预览模式的默认语言行为

当前不覆盖：

- AI 自由回复内容的强制翻译
- workbook 中的业务数据、项目名、字段名
- `zh` / `en` 之外的第三种界面语言

## 2. 语言来源与解析规则

当前前端不自行探测浏览器或系统语言，而是依赖宿主 `bridge.getHostContext` 返回：

- `resolvedUiLocale`
- `uiLanguageOverride`

当前解析规则：

- 当 `uiLanguageOverride = zh` 时，任务窗格固定 UI 直接显示中文
- 当 `uiLanguageOverride = en` 时，任务窗格固定 UI 直接显示英文
- 当 `uiLanguageOverride = system` 时：
  - Excel UI 语言属于 `zh-*`，则 `resolvedUiLocale = zh`
  - 其他所有 Excel UI 语言都归一到 `resolvedUiLocale = en`
- 如果宿主上下文读取失败，前端首屏和后续固定 UI 都回退到英文，避免英文 Excel 下先闪中文

这意味着当前双语规则是“仅中文 Excel 显示中文，其余全部英文”。

## 3. 宿主驱动的语言切换行为

任务窗格语言当前由 Excel 宿主驱动，不是由用户输入语言驱动。

当前可见行为：

- 在中文 Excel 中打开任务窗格，固定 UI 与前端系统消息显示中文
- 在英文 Excel 中打开任务窗格，固定 UI 与前端系统消息显示英文
- 同一份前端包无需切换构建产物，依赖宿主返回的 `resolvedUiLocale` 决定显示语言
- 当前不要求 Excel 运行过程中热切换 Office UI 语言后，已打开任务窗格立即自动刷新；重新打开任务窗格后应按最新宿主语言重新初始化

## 4. 已本地化的任务窗格范围

当前已本地化的前端固定 UI 包括：

- 页眉、会话列表按钮、设置按钮、发送按钮
- 会话抽屉、会话占位标题（如 `未命名会话` / `Untitled`）
- 设置对话框、按钮文案、加载/失败提示；设置对话框在较矮任务窗格中限制在视口内并允许内部纵向滚动，确保底部保存按钮可达
- 选区胶囊的空状态与格式化摘要
- 欢迎语与初始线程系统消息

当前已本地化的前端系统消息包括：

- Excel 命令确认卡片标题、摘要、详情、按钮
- `upload_data` 预览确认卡片与本地取消提示
- Planner 计划预览标题、步骤格式化、执行日志状态标签
- 前端本地生成的常规兜底错误、加载失败、删除确认等消息

边界说明：

- 宿主生成的错误或状态消息由宿主侧先完成本地化，再传给前端显示
- 前端不会二次翻译宿主返回消息
- backend 或 AI 直接返回的自由文本（例如部分 assistant message、plan summary、业务返回 message）仍按原样显示，不强制改写语言
- 浏览器预览内部少量校验异常和开发态 mock error 当前仍以英文实现为主，不属于已承诺完成的双语覆盖面

设置对话框当前持久化这些 LLM / 业务连接配置：

- `API Key`
- `Base URL`：普通 Planner 与 AI 列映射共用的模型 endpoint
- `Business Base URL`：业务 API 或 mock server endpoint
- `Model`
- `API Format`：`OpenAI Compatible` 或 `Anthropic Messages`
- `SSO URL`
- `Login success path`

普通 Planner 会复用 `Base URL`、`API Key`、`Model` 和 `API Format`。`OpenAI Compatible` 使用 chat completions 请求；`Anthropic Messages` 使用 Anthropic Messages 请求，并把 system prompt 放在顶层 `system` 字段。

## 5. 浏览器预览模式

浏览器预览模式下没有真实 Excel 宿主，当前默认行为是：

- `resolvedUiLocale` 默认返回 `en`
- `uiLanguageOverride` 默认返回 `system`
- 因此直接在浏览器预览打开前端时，固定 UI 默认显示英文

如果在浏览器预览里保存设置并把 `uiLanguageOverride` 改成 `zh` 或 `en`：

- `getHostContext` 会回传新的覆盖值
- 前端会重新读取宿主上下文，固定 UI 会按该覆盖值切换

补充说明：

- 当前设置界面还没有暴露手动切换语言的正式入口
- 但浏览器预览和宿主协议已经保留 `uiLanguageOverride`，用于后续接入手动切换能力

当前浏览器预览切换的是外围固定 UI 和 mock 系统文案；示例 workbook 名、selection 样例数据本身仍是固定 demo 内容，不随 `uiLanguageOverride` 翻译。

## 6. AI 回复与语言策略

当前双语切换只约束固定 UI 和系统消息，不强制 AI 自由回复跟随 Excel UI 语言。

当前策略：

- free-form chat / planner assistant message 会尽量跟随用户输入语言
- 英文 Excel 中输入中文问题时，AI 仍应尽量用中文回答
- 中文 Excel 中输入英文问题时，AI 仍可尽量用英文回答

因此会出现“英文固定 UI + 中文 AI 回复”或“中文固定 UI + 英文 AI 回复”的组合，这属于当前预期行为，不视为本地化缺陷。

## 7. 主要代码入口

如果后续继续迭代 Task Pane 双语行为，建议优先看：

- 前端壳层与消息渲染
  - [src/OfficeAgent.Frontend/src/App.tsx](../../src/OfficeAgent.Frontend/src/App.tsx)
- 前端语言包
  - [src/OfficeAgent.Frontend/src/i18n/uiStrings.ts](../../src/OfficeAgent.Frontend/src/i18n/uiStrings.ts)
- bridge 协议与浏览器预览 fallback
  - [src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts](../../src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts)
- 宿主语言解析与 bridge 上下文
  - [src/OfficeAgent.ExcelAddIn/Localization/UiLocaleResolver.cs](../../src/OfficeAgent.ExcelAddIn/Localization/UiLocaleResolver.cs)
  - [src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs](../../src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs)

## 8. 主要测试入口

- 前端固定 UI / 系统消息
  - [src/OfficeAgent.Frontend/src/App.test.tsx](../../src/OfficeAgent.Frontend/src/App.test.tsx)
- 浏览器预览 locale fallback
  - [src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts](../../src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts)
- 宿主 locale 归一规则
  - [tests/OfficeAgent.ExcelAddIn.Tests/UiLocaleResolverTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/UiLocaleResolverTests.cs)
- 宿主双语字符串
  - [tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/HostLocalizedStringsTests.cs)
- bridge 宿主上下文协议
  - [tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs](../../tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs)

## 9. 相关文档

- 双语设计说明
  - [docs/superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md](../superpowers/specs/2026-04-24-office-agent-bilingual-ui-design.md)
- 手工测试
  - [docs/vsto-manual-test-checklist.md](../vsto-manual-test-checklist.md)
- Ribbon Sync 快照
  - [docs/modules/ribbon-sync-current-behavior.md](./ribbon-sync-current-behavior.md)

## 10. 文档维护约定

如果任务窗格用户可见行为发生变化，至少同步更新：

- 本文第 2 节到第 6 节
- [docs/module-index.md](../module-index.md)
- [docs/vsto-manual-test-checklist.md](../vsto-manual-test-checklist.md)
