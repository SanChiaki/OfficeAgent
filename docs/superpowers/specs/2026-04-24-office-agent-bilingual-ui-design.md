# OfficeAgent Excel UI 双语切换设计

日期：2026-04-24

状态：设计已确认，待实施

## 1. 目标

为 OfficeAgent 增加中英文 UI 切换能力，并满足以下目标：

- 默认根据 Excel UI 语言自动切换
- 仅当 Excel UI 语言属于 `zh-*` 时显示中文 UI
- 其他所有语言一律显示英文 UI
- 覆盖任务窗格 React UI、Ribbon、WinForms 对话框、`MessageBox`、bridge 宿主消息、前端生成的系统消息
- 不改变 AI 自由回复策略，AI 仍尽量跟随用户输入语言
- 设计上预留未来“手动切换插件语言”的能力，但本次不要求必须暴露设置入口

## 2. 背景

当前插件的用户可见文案主要是硬编码中文，且分散在多个层面：

- `src/OfficeAgent.Frontend/src/App.tsx`
- `src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx`
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- `src/OfficeAgent.ExcelAddIn/Dialogs/*.cs`
- `src/OfficeAgent.ExcelAddIn/SsoLoginPopup.cs`
- `src/OfficeAgent.ExcelAddIn/WebBridge/*.cs`

这带来几个直接问题：

- 英文 Excel 环境下仍显示中文 UI
- 任务窗格、Ribbon、原生弹窗各自维护文案，未来容易切换不一致
- 前端没有语言来源，只能写死默认语言
- 若后续增加手动语言切换，当前结构缺少统一入口

因此需要把“UI 语言来源”和“文案提供方式”从零散硬编码改成明确的双语架构。

## 3. 范围

### 3.1 本次要做

- 引入统一的 UI 语言解析模型，只支持 `zh` / `en`
- 宿主侧根据 Excel UI 语言解析当前语言
- 在 `AppSettings` 中预留手动覆盖字段
- bridge 增加只读宿主上下文接口，把解析后的语言传给前端
- C# 侧集中接管 Ribbon、WinForms、`MessageBox`、bridge 宿主消息的中英文文案
- React 侧集中接管固定 UI 和前端生成系统消息的中英文文案
- 为浏览器预览模式补齐同样的双语逻辑
- 补齐相关单元测试和手工验证说明

### 3.2 本次不做

- 不支持 `zh` / `en` 之外的第三种语言
- 不引入重量级 i18n 框架
- 不强制 AI 自由回复跟随 Excel UI 语言
- 不翻译 workbook 中的业务数据、项目名、字段名或业务系统返回值
- 不要求 Excel 运行过程中切换 Office UI 语言后立即热更新所有已打开窗口

## 4. 语言规则

插件内部只认一个最终结果：`resolvedUiLocale`。

- 允许值：`zh`、`en`
- 解析优先级：
  1. 用户手动覆盖语言
  2. Excel UI 语言自动检测
  3. 检测失败时回退到 `en`

手动覆盖字段建议命名为 `uiLanguageOverride`，允许值：

- `system`
- `zh`
- `en`

当前版本即使不暴露手动切换入口，也要求底层支持该字段，默认值为 `system`。

自动检测规则固定为：

- 只要 Excel UI 语言表示 `zh-*`，则 `resolvedUiLocale = zh`
- 其他任何语言都归为 `resolvedUiLocale = en`

这保证当前需求“仅在中文语言下展示中文 UI，其余都是英文 UI”被严格表达为一个稳定规则。

## 5. 方案对比

### 5.1 方案一：统一语言源，前后端各自本地化

做法：

- 宿主侧统一解析语言
- 前端通过 bridge 获取 `resolvedUiLocale`
- C# 和 React 各自维护轻量双语文案层

优点：

- 任务窗格、Ribbon、WinForms、宿主消息能稳定保持一致
- 语言判断只有一处，便于后续增加手动切换
- 不需要引入复杂工具链

缺点：

- 初次改动面较大，需要迁移零散硬编码

### 5.2 方案二：只给前端加 i18n，C# 侧继续点状判断

优点：

- 任务窗格改动快

缺点：

- Ribbon、MessageBox、对话框、bridge 宿主消息依然分散
- 极易出现前端已英文但宿主仍中文的裂缝
- 后续手动切换会继续扩散条件判断

### 5.3 方案三：依赖 `.resx` 或线程文化做全局切换

优点：

- 接近传统 WinForms/VSTO 本地化方式

缺点：

- 当前系统是 VSTO + WebView2 + React 的跨端结构
- 前端无法自然复用 `.resx`
- 产品需求绑定的是 Excel UI 语言，而不是简单依赖 `.NET CurrentUICulture`

### 5.4 结论

采用方案一：统一语言源，前后端各自实现轻量本地化。

## 6. 推荐设计

### 6.1 统一语言解析层

在 Excel add-in 宿主侧新增一个统一的 UI 语言解析组件，例如：

- `UiLocaleResolver`
- `IUiLocaleProvider`

它负责：

- 读取持久化设置中的 `uiLanguageOverride`
- 读取 Excel UI 语言
- 输出 `resolvedUiLocale`

推荐行为：

- `uiLanguageOverride = zh` 时直接返回 `zh`
- `uiLanguageOverride = en` 时直接返回 `en`
- `uiLanguageOverride = system` 时按 Excel UI 语言归一
- 任何异常场景回退到 `en`

Excel UI 语言的具体读取方式建议使用 Office/Excel 提供的 UI language API，例如通过 `Application.LanguageSettings` 获取 UI 语言标识，再归一到 `zh` / `en`。

### 6.2 设置模型预留

在 `AppSettings` 中新增字段：

```csharp
public string UiLanguageOverride { get; set; } = "system";
```

`FileSettingsStore` 需要同步支持：

- 加载旧设置文件时缺省为 `system`
- 保存时原样持久化
- 若值非法，则按 `system` 处理

这里持久化的是覆盖值，不是最终计算结果：

- `uiLanguageOverride` 持久化
- `resolvedUiLocale` 不持久化

### 6.3 bridge 宿主上下文

新增只读 bridge 消息：

- `bridge.getHostContext`

返回 payload 建议至少包含：

```json
{
  "resolvedUiLocale": "zh",
  "uiLanguageOverride": "system"
}
```

目的：

- 给前端一个显式宿主上下文
- 避免把 UI 运行时信息塞进业务设置对象
- 为未来设置页加入手动语言切换保留协议位

前端不自行探测语言，只消费该桥接结果。

### 6.4 C# 侧文案组织

C# 宿主侧负责所有“宿主自己生成”的用户可见文案，包括：

- Ribbon tab/group/button/dropdown label
- Ribbon 下拉占位和状态文字
- `MessageBox` 标题和正文
- `TaskPaneHostControl` 中 WebView2 缺失或初始化失败时的宿主提示
- Ribbon Sync 确认与结果提示
- `ProjectLayoutDialog`
- `DownloadConfirmDialog`
- `UploadConfirmDialog`
- `OperationResultDialog`
- `SsoLoginPopup`
- WebBridge 错误消息、登录结果消息、fallback HTML

这里不建议继续散落字符串，而是集中到一个本地化文案提供层，例如：

- `UiText`
- `HostLocalizedStrings`

推荐采用“按场景方法”而不是纯 key-value 字典，例如：

- `SelectProjectPlaceholder()`
- `ProjectLoadFailed(string message)`
- `ConfirmDownload(string operationName, string projectName, int rowCount, int fieldCount, SyncOperationPreview preview)`
- `ProjectLayoutDialogTitle()`

这样更适合当前大量带参数拼接的消息模板。

### 6.5 React 侧文案组织

React 前端负责所有“前端自己生成”的固定 UI 和系统消息，包括：

- 顶部按钮、抽屉、设置对话框、删除确认框
- 欢迎语
- 确认卡片标题和按钮
- 前端本地产生的取消消息
- 前端本地失败兜底消息
- 计划摘要标题和计划步骤格式化
- 浏览器预览模式文案

不建议引入 `react-intl` 之类重量级方案，当前体量更适合新增一个轻量模块，例如：

- `src/OfficeAgent.Frontend/src/i18n/uiStrings.ts`

结构建议：

- 导出 `zh` / `en` 两套字典
- 提供 `getUiStrings(locale)` 返回当前语言包
- 对带参数的文案使用函数值，例如：
  - `requestFailed(message)`
  - `deleteSessionPrompt(title)`
  - `bridgeConnected(host, version)`

### 6.6 谁生成消息，谁负责本地化

本次设计遵循一个明确规则：

- 宿主生成的消息，在宿主侧本地化
- 前端生成的消息，在前端侧本地化

这意味着：

- `bridge` 返回给前端的错误或状态消息应已经是最终字符串
- 前端不再试图重拼宿主错误
- 前端内部即时生成的消息，如欢迎语、确认取消提示、计划标题，仍由前端本地化

该规则能降低跨端拼接导致的错位和重复逻辑。

### 6.7 前端启动与刷新行为

前端启动时除现有的：

- `getSessions`
- `getSettings`
- `getSelectionContext`
- `getLoginStatus`

之外，再调用：

- `getHostContext`

前端行为：

- 成功时用 `resolvedUiLocale` 驱动全量 UI 文案
- 失败时回退到 `en`

前端初始语言也建议默认 `en`，避免英文 Excel 下首屏短暂闪现中文。

### 6.8 Ribbon 和原生窗口刷新语义

Ribbon 和原生窗口采用“打开时取值、切换时显式刷新”的策略：

- Add-in 启动时初始化 Ribbon 文案
- 打开对话框时按当前 `resolvedUiLocale` 取文案
- 未来若加入手动切换，保存后触发：
  - Ribbon 重新设置 `Label` 并 `Invalidate`
  - 新打开的对话框直接使用新语言
  - 前端重新拉取或本地刷新 `hostContext`

当前不要求支持运行中的 Excel UI 语言热切换自动刷新所有已打开窗口。

## 7. 具体协议与模型建议

### 7.1 `AppSettings`

新增字段：

```json
{
  "uiLanguageOverride": "system"
}
```

允许值仅有：

- `system`
- `zh`
- `en`

### 7.2 bridge 新消息

新增常量：

- `bridge.getHostContext`

新增前端类型：

```ts
export interface HostContext {
  resolvedUiLocale: 'zh' | 'en';
  uiLanguageOverride: 'system' | 'zh' | 'en';
}
```

`nativeBridge` 新增：

- `getHostContext()`

浏览器预览模式默认返回：

```json
{
  "resolvedUiLocale": "en",
  "uiLanguageOverride": "system"
}
```

### 7.3 文案层命名约束

为避免未来跨端继续分裂，本次要求：

- 前端和宿主的语言值都只使用 `zh` / `en`
- 前端和宿主使用相同的 override 枚举：`system | zh | en`
- 新增用户可见文案时，禁止继续直接写死在业务流程中

## 8. 任务分解建议

### 8.1 阶段一：语言源与协议

- 在 `AppSettings` 和 `FileSettingsStore` 中增加 `uiLanguageOverride`
- 新增宿主语言解析器
- `bridge` 增加 `getHostContext`
- 补协议与模型测试

### 8.2 阶段二：宿主 UI 本地化

- Ribbon labels 全面改为集中式取文案
- WinForms / `MessageBox` / `SsoLoginPopup` 改为集中式取文案
- `WebMessageRouter` / `WebViewBootstrapper` 宿主消息改为集中式取文案

### 8.3 阶段三：前端 UI 本地化

- React 固定 UI 改为 `uiStrings`
- 欢迎语、取消提示、计划标题、步骤格式化改为 `uiStrings`
- 浏览器预览 mock 文案同步改造

### 8.4 阶段四：测试与文档

- 补齐双语分支测试
- 更新手工测试清单
- 更新模块快照文档

## 9. 测试策略

### 9.1 前端测试

Vitest 至少覆盖：

- `zh` 宿主上下文时显示中文固定 UI
- `en` 宿主上下文时显示英文固定 UI
- 欢迎语、确认卡片、删除确认、取消提示、计划标题、浏览器预览文案随语言切换
- `getHostContext` 失败时默认英文

### 9.2 Add-in / bridge 测试

至少覆盖：

- `uiLanguageOverride` 默认值和持久化
- `resolvedUiLocale` 的归一规则
- `bridge.getHostContext` 返回结构
- `bridge` 宿主错误消息在 `zh` / `en` 下均正确

### 9.3 Ribbon / Dialog 测试

至少覆盖：

- Ribbon labels
- 项目下拉框占位和状态文案
- `ProjectLayoutDialog`
- `DownloadConfirmDialog`
- `UploadConfirmDialog`
- `OperationResultDialog`
- `SsoLoginPopup`

### 9.4 手工验证

至少验证两个环境：

- 中文 Excel
- 英文 Excel

场景包括：

- 打开任务窗格首屏
- 查看 Ribbon 文案
- 打开登录弹窗
- 触发 Ribbon Sync 确认和结果提示
- 在英文 Excel 下输入中文问题，确认 AI 自由回复仍可按用户输入语言输出，而不是被强制改成英文

## 10. 风险与对策

### 10.1 文案散落导致遗漏

风险：

- 某些用户可见消息仍留在原业务代码里，双语切换不完整

对策：

- 实施前先清点用户可见文案
- 实施时按“入口面”而不是按文件数量验收

### 10.2 Designer 默认值残留中文

风险：

- `AgentRibbon.Designer.cs` 初始中文在英文环境下短暂或永久残留

对策：

- Add-in 启动时主动刷新所有 Ribbon label
- 不依赖 Designer 默认值作为最终显示结果

### 10.3 前端首屏闪中文

风险：

- 当前前端初始 state 带中文硬编码

对策：

- 语言未就绪前默认英文
- 首屏展示全部走 `uiStrings`

### 10.4 未来手动切换需要推翻本次架构

风险：

- 当前若只做自动检测，后面再加 override 会重做协议和存储

对策：

- 本次就引入 `uiLanguageOverride` + `resolvedUiLocale` 双层模型

## 11. 文档更新要求

由于本次改动会改变 Ribbon Sync 的用户可见行为，实施时至少同步更新：

- `docs/modules/ribbon-sync-current-behavior.md`
- `docs/vsto-manual-test-checklist.md`
- `docs/modules/task-pane-current-behavior.md`
- `docs/module-index.md`

其中：

- `ribbon-sync-current-behavior.md` 需要说明 Ribbon、登录弹窗、Ribbon Sync 原生提示的双语行为
- `task-pane-current-behavior.md` 需要作为新模块快照，说明任务窗格固定 UI、前端系统消息和浏览器预览模式的语言切换行为

## 12. 结论

本次应以“宿主统一解析语言，前后端各自集中本地化”为核心方案落地双语 UI。

该方案能在不改变 AI 自由回复策略的前提下，统一覆盖：

- 任务窗格固定 UI
- 前端生成的系统消息
- Ribbon
- WinForms 对话框
- `MessageBox`
- bridge 宿主消息

同时通过 `uiLanguageOverride + resolvedUiLocale` 双层模型，为未来“手动切换插件语言”保留清晰且低成本的扩展路径。
