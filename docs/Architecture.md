# OfficeAgent 架构文档

## 1. 系统概述

OfficeAgent 是一个运行在 Excel 进程内的 AI Agent，通过侧边任务窗格（Task Pane）提供自然语言交互能力。用户可以用中文或英文描述需求，Agent 会规划并执行 Excel 操作（读写单元格、管理工作表）或业务技能（上传数据到外部系统）。

```
┌─────────────────────────────────────────────────────────┐
│                      Excel 进程 (STA)                    │
│                                                          │
│  ┌──────────────┐    ┌───────────────────────────────┐   │
│  │ AgentRibbon  │    │        TaskPaneHostControl     │   │
│  │ (切换面板)    │───▶│  ┌─────────────────────────┐  │   │
│  └──────────────┘    │  │     WebView2 控件         │  │   │
│                      │  │  ┌─────────────────────┐  │  │   │
│                      │  │  │  React/TS 前端       │  │  │   │
│                      │  │  │  (消息列表 + 输入框)  │  │  │   │
│                      │  │  └─────────────────────┘  │  │   │
│                      │  └─────────────────────────┘  │   │
│                      └───────────────────────────────┘   │
│                              │ bridge 消息               │
│                              ▼                           │
│  ┌──────────────────────────────────────────────────┐    │
│  │              WebMessageRouter                     │    │
│  │         (消息路由 + 确认服务)                      │    │
│  └───────────┬──────────────────┬───────────────────┘    │
│              │                  │                         │
│              ▼                  ▼                         │
│  ┌──────────────────┐  ┌──────────────────────┐          │
│  │ ExcelInteropAdapter │  │  AgentOrchestrator   │          │
│  │ (Excel COM 操作)   │  │  (Agent 调度循环)     │          │
│  └──────────────────┘  └──────┬───────────────┘          │
│                                │                          │
│                 ┌──────────────┼──────────────┐           │
│                 ▼              ▼              ▼           │
│          LlmPlannerClient  PlanExecutor  SkillRegistry   │
│          (LLM API 调用)    (计划执行)      (技能注册)      │
└─────────────────────────────────────────────────────────┘
```

## 2. 项目分层

```
src/
├── OfficeAgent.ExcelAddIn/   ← VSTO 宿主层（COM 互操作）
├── OfficeAgent.Core/         ← 领域层（纯逻辑，无 COM 依赖）
├── OfficeAgent.Infrastructure/ ← 基础设施层（IO、HTTP、加密）
└── OfficeAgent.Frontend/     ← 前端层（React + TypeScript）
```

| 层 | 项目 | 职责 | 关键类型 |
|---|---|---|---|
| 宿主 | `ExcelAddIn` | VSTO 入口、Ribbon、任务窗格生命周期、WebView2 启动、Excel 事件桥接、COM 互操作 | `ThisAddIn`, `WebViewBootstrapper`, `WebMessageRouter`, `ExcelInteropAdapter` |
| 领域 | `Core` | 领域模型、Agent 调度、技能注册、确认服务、计划执行 | `AgentOrchestrator`, `PlanExecutor`, `SkillRegistry`, `ConfirmationService` |
| 基础设施 | `Infrastructure` | HTTP 客户端（LLM、业务 API）、文件存储、DPAPI 加密、日志 | `LlmPlannerClient`, `BusinessApiClient`, `FileSessionStore`, `FileSettingsStore` |
| 前端 | `Frontend` | 聊天 UI、会话管理、设置面板、桥接客户端 | `App`, `NativeBridge` |

**依赖方向：** `ExcelAddIn` → `Core` → `Infrastructure`。Core 不引用 ExcelAddIn，Infrastructure 不引用 ExcelAddIn。前端通过 WebView2 桥接与后端通信，无直接依赖。

## 3. 启动与组合根

`ThisAddIn_Startup` 是整个应用的组合根，按顺序创建并注入所有依赖：

```
FileLogSink + OfficeAgentLog.Configure()          ← 日志
FileSessionStore                                   ← 会话持久化
FileSettingsStore + DpapiSecretProtector            ← 设置（API Key 加密）
ExcelSelectionContextService(Application)           ← 选区读取
ExcelInteropAdapter(Application, ...)               ← Excel 命令执行
ExcelFocusCoordinator(Application)                  ← 焦点协调
SkillRegistry(UploadDataSkill(...))                 ← 技能注册
AgentOrchestrator(skillRegistry, selectionService,  ← Agent 调度器
  excelCommandExecutor, llmPlannerClient, planExecutor)
TaskPaneController(this, orchestrator, sessionStore, ← 任务窗格管理
  settingsStore, selectionService, focusCoordinator)
```

启动完成后，订阅 `Application.SheetSelectionChange` 将选区上下文推送到前端。

## 4. WebView2 桥接协议

### 4.1 消息格式

前端与宿主之间通过 JSON 消息通信，使用 WebView2 的 `postMessage` / `WebMessageReceived` 机制。

**请求-响应模式：**

```jsonc
// 请求
{ "type": "bridge.<action>", "requestId": "<uuid>", "payload": <data> }

// 成功响应
{ "type": "bridge.<action>", "requestId": "<uuid>", "ok": true, "payload": <data> }

// 失败响应
{ "type": "bridge.<action>", "requestId": "<uuid>", "ok": false, "error": { "code": "...", "message": "..." } }
```

**事件推送模式（宿主→前端，无 requestId）：**

```json
{ "type": "bridge.selectionContextChanged", "payload": <SelectionContext> }
```

### 4.2 桥接类型

| 类型 | 方向 | 说明 |
|------|------|------|
| `bridge.ping` | FE→Host | 心跳检测 |
| `bridge.getSettings` | FE→Host | 获取设置 |
| `bridge.saveSettings` | FE→Host | 保存设置 |
| `bridge.getSessions` | FE→Host | 获取会话列表 |
| `bridge.getSelectionContext` | FE→Host | 获取当前选区 |
| `bridge.executeExcelCommand` | FE→Host | 执行 Excel 命令（直接） |
| `bridge.runSkill` | FE→Host | 执行技能（上传等） |
| `bridge.runAgent` | FE→Host | Agent 规划 |
| `bridge.saveSessions` | FE→Host | 持久化会话列表 |
| `bridge.selectionContextChanged` | Host→FE | 选区变化事件 |

### 4.3 前端桥接客户端

`NativeBridge` 通过 `requestId` 关联请求和响应。每次 `invoke()` 生成 UUID，将 `{resolve, reject}` 存入 `pendingRequests` Map，响应到达后匹配并 resolve Promise。

当 `window.chrome.webview` 不存在时（浏览器开发模式），所有操作返回 mock 数据，支持脱离 Excel 进行前端开发。

## 5. Agent 调度流程

### 5.1 输入分类

前端 `handleComposerSend` 对用户输入做三层分类：

```
用户输入
  │
  ├─ 匹配 Excel 直接命令（正则）
  │   /readSelection, /addSheet, /writeRange, /renameSheet, /deleteSheet
  │   → nativeBridge.executeExcelCommand()
  │
  ├─ 匹配技能命令
  │   /upload_data, "上传到...", "upload ... to ..."
  │   → nativeBridge.runSkill()
  │
  └─ 其他
      → nativeBridge.runAgent()
```

### 5.2 Agent 循环

```
                    AgentCommandEnvelope
                    { userInput, sessionId, conversationHistory }
                              │
                              ▼
                    ┌─ Confirmed + Plan? ──┐
                    │ YES                   │ NO
                    ▼                       ▼
              ExecuteFrozenPlan()     Build PlannerRequest
              (PlanExecutor)          { userInput, SelectionContext,
                                       ConversationHistory }
                                              │
                                              ▼
                                     ┌── 重试循环 (max 3) ──┐
                                     │                       │
                                     │   LlmPlannerClient    │
                                     │   POST /chat/completions
                                     │         │             │
                                     │         ▼             │
                                     │   PlannerResponse     │
                                     │         │             │
                                     │   ┌─ mode? ──────────┤
                                     │   │                  │
                                     │   ├─ "message"       │
                                     │   │  → 返回聊天回复    │
                                     │   │                  │
                                     │   ├─ "read_step"     │
                                     │   │  → 执行 Excel 读取 │
                                     │   │  → 追加 Observation│
                                     │   │  → 继续循环        │──┘
                                     │   │
                                     │   └─ "plan"
                                     │      → 返回待确认计划
                                     │
                                     └── 3 次用尽 → 返回失败
```

### 5.3 对话历史注入

每次 `bridge.runAgent` 调用时，前端提取最近 10 轮对话（20 条消息，仅 `user`/`assistant` 角色），作为 `conversationHistory` 传入。LLM 收到的消息数组结构为：

```
[system: planner 指令] → [history: prior turns] → [user: 当前请求]
```

数据流路径：`AgentCommandEnvelope.ConversationHistory` → `PlannerRequest.ConversationHistory` → `LlmPlannerClient.BuildChatMessages()`。

### 5.4 确认流程

所有写操作（Excel 命令、技能、Agent 计划）采用两阶段确认：

1. **预览阶段**：后端返回 `requiresConfirmation: true` + 预览信息
2. **前端展示**：渲染 `ConfirmationCard`，等待用户操作
3. **确认执行**：用户点击确认后，前端重新发送请求（`confirmed: true` + 冻结的计划/命令）

### 5.5 计划执行

`PlanExecutor` 按顺序执行计划步骤：
- Excel 步骤 → `ConfirmationService.Validate()` → `IExcelCommandExecutor.Execute()`
- 上传步骤 → `UploadDataSkill`（预览→执行）
- 首次失败后停止，剩余步骤标记为 `skipped`

## 6. Excel COM 互操作

### 6.1 线程模型

Excel 是 STA（单线程单元）进程。所有 COM 对象只能在创建它们的 UI 线程上访问。

```
UI/STA 线程                    线程池
    │
    ├── COM 调用（读选区、写单元格等）
    │
    ├── await LlmPlannerClient.CompleteAsync()
    │       │                     ──────▶  HTTP 请求在线程池执行
    │       │                                     ConfigureAwait(false)
    │       ◀────── 响应返回 ──────
    │
    ├── ConfigureAwait(true) 确保续体回到 STA 线程
    │
    └── 后续 COM 调用安全
```

关键措施：
- `WebViewBootstrapper.InitializeAsync` 中显式安装 `WindowsFormsSynchronizationContext`
- Orchestrator 层使用 `ConfigureAwait(true)` 回到 UI 线程
- HTTP 层使用 `ConfigureAwait(false)` 不占用 UI 线程
- `isProcessing` 标志保证同一时间只有一个 Agent 请求在执行

### 6.2 Excel 命令

| 命令 | 说明 | 需要确认 |
|------|------|---------|
| `readSelectionTable` | 读取选区为表格（表头+数据行） | 否 |
| `writeRange` | 写入二维数组到目标地址 | 是 |
| `addWorksheet` | 在末尾添加工作表 | 是 |
| `renameWorksheet` | 重命名工作表 | 是 |
| `deleteWorksheet` | 删除工作表（抑制 Excel 确认弹窗） | 是 |

### 6.3 安全守卫

`ConfirmationService.Validate()` 在执行前检查：
- 工作表是否受保护（`ProtectContents`）
- 工作簿结构是否受保护（`ProtectStructure`）
- 目标区域是否包含合并单元格
- 写入数据维度是否合规

`ExcelOperationGuard` 提供底层检查方法，所有写操作执行前必须通过。

## 7. 前端状态管理

前端使用 React `useState` 管理所有状态，无外部状态库。

```
sessions: ChatSession[]                    ← 会话列表
activeSessionId: string                    ← 当前活跃会话
sessionThreads: Record<sessionId, ThreadMessage[]>  ← 每会话消息列表
pendingConfirmations: Record<sessionId, PendingConfirmation>  ← 待确认操作
pendingCommandSessions: Record<sessionId, boolean>  ← 加载状态
selectionContext: SelectionContext | null   ← Excel 选区（宿主推送）
settings / draftSettings                   ← 设置管理
```

`PendingConfirmation` 分三种类型：
- `excel` — 待确认的 Excel 命令 + 预览
- `skill` — 待确认的上传操作 + 预览
- `agent` — 待确认的 Agent 计划 + 预览

## 8. LLM 调用

### 8.1 请求构造

`LlmPlannerClient` 向 OpenAI 兼容的 `/v1/chat/completions` 端点发送请求：

```jsonc
{
  "model": "<settings.Model>",
  "messages": [
    { "role": "system", "content": "<planner 指令>" },
    // ... conversationHistory (prior turns)
    { "role": "user", "content": "Planner request:\n<PlannerRequest JSON>" }
  ],
  "response_format": { "type": "json_object" }
}
```

### 8.2 响应格式

LLM 被要求返回固定结构的 JSON：

```jsonc
{
  "mode": "message | read_step | plan",
  "assistantMessage": "自然语言回复",
  "step": null | { "type": "...", "args": {} },
  "plan": null | { "summary": "...", "steps": [...] }
}
```

### 8.3 降级策略

如果主端点返回 404/405，自动降级到旧版 `/planner` 端点（`{ model, request }` 格式）。

## 9. 存储与安全

| 数据 | 文件位置 | 加密 |
|------|---------|------|
| 会话 | `%LocalAppData%\OfficeAgent\sessions\sessions.json` | 无 |
| 设置 | `%LocalAppData%\OfficeAgent\settings.json` | API Key 用 DPAPI 加密 |
| 日志 | `%LocalAppData%\OfficeAgent\logs\officeagent.log` | 无 |

- DPAPI 使用 `DataProtectionScope.CurrentUser`，密钥绑定 Windows 用户账户
- 日志格式为 JSON Lines（每行一个 JSON 对象）
- 设置加载失败（加密损坏、格式错误）时回退为空默认值，不崩溃

## 10. MSI 安装包

### 10.1 安装包结构

使用 WiX v4 构建，per-user 范围，安装到 `%LocalAppData%\OfficeAgent\ExcelAddIn\`。

**前置条件检查：**
- VSTO Runtime 4.0（注册表搜索 HKLM）
- WebView2 Runtime（注册表搜索 HKLM/HKCU）

**部署内容：**
- VSTO 运行时工具 DLL
- WebView2 核心/WinForms/Wpf DLL + 原生 Loader（arm64/x64/x86）
- Newtonsoft.Json
- OfficeAgent.Core.dll / OfficeAgent.Infrastructure.dll / OfficeAgent.ExcelAddIn.dll
- VSTO 清单文件（.vsto + .dll.manifest）
- 前端资源（frontend/index.html + frontend/assets/）

**注册表写入：**
`HKCU\Software\Microsoft\Office\Excel\Addins\OfficeAgent.ExcelAddIn`
- `LoadBehavior = 3`（启动时加载）
- `Manifest = file:///[#filOfficeAgentExcelAddInVsto]|vstolocal`

### 10.2 版本号

从 `git rev-list --count HEAD` 自动生成，格式 `1.0.<commitCount>`（如 `1.0.35`）。`MajorUpgrade` 使用固定 `UpgradeCode` 支持覆盖安装。

### 10.3 CI 构建

GitHub Actions workflow（`.github/workflows/build-msi.yml`）在 `windows-latest` 上构建。CI 环境通过 `New-SelfSignedCertificate` 生成临时代码签名证书解决 VSTO 清单签名要求。

## 11. 领域模型

```
AgentCommandEnvelope                    AgentCommandResult
├── DispatchMode: auto|agent|skill      ├── Route: chat|excelCommand|skill|plan
├── SessionId                           ├── Status: completed|failed|preview
├── UserInput                           ├── Message
├── Confirmed                           ├── RequiresConfirmation
├── Plan (AgentPlan)                    ├── Planner (PlannerResponse)
├── ConversationHistory (ConversationTurn[])  ├── Journal (PlanExecutionJournal)
└── UploadPreview                       └── Preview / UploadPreview

PlannerRequest                          PlannerResponse
├── SessionId                           ├── Mode: message|read_step|plan
├── UserInput                           ├── AssistantMessage
├── SelectionContext                    ├── Step (PlannerStep)
├── Observations (PlannerObservation[]) └── Plan (AgentPlan)
└── ConversationHistory (ConversationTurn[])

ConversationTurn                        AgentPlan
├── Role: user|assistant                ├── Summary
└── Content                             └── Steps (AgentPlanStep[])
                                              ├── Type
                                              └── Args (JObject)
```

所有模型使用 `[JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]`，序列化为 camelCase JSON。
