# Analytics 真实埋点系统对接指南

本文说明 OfficeAgent / Resy AI 如何从本地 mock server 切换到真实内网埋点系统。

适用范围：

- Ribbon、Ribbon Sync、配置模板和任务窗格交互埋点
- `SystemConnector` / 业务连接器内部扩展埋点
- 内网接口 `POST https://{HOST}/insertLog`

不适用范围：

- 诊断日志 `%LocalAppData%\OfficeAgent\logs\officeagent.log`
- 业务接口 `/projects`、`/head`、`/find`、`/batchSave`
- 大模型接口 `Base URL`

## 1. 对接结论

真实系统只需要提供一个接口：

```http
POST https://{HOST}/insertLog
Content-Type: application/json
```

插件隐藏配置中保存完整埋点 URL：

```text
AnalyticsUrl = https://{HOST}/insertLog
```

当前不在任务窗格 Settings UI 中展示该配置。插件会直接请求 `AnalyticsUrl`，不会再自动拼接 `/insertLog`。

本地配置文件路径为 `%LocalAppData%\OfficeAgent\settings.json`。可在该 JSON 中加入或更新：

```json
{
  "AnalyticsUrl": "https://{HOST}/insertLog"
}
```

保留文件中的其他字段不变。修改后需要重启 Excel。

如果内网系统路径带前缀，例如：

```text
https://analytics.internal.example/logging/insertLog
```

就把完整地址保存为 `AnalyticsUrl`。不要只保存域名或基地址。

## 2. 外层接口合同

插件发出的外层 JSON 固定为：

```json
{
  "frontEndIntent": "excelAi",
  "clientSource": "Excel",
  "questionType": 1,
  "askId": "zzZMg0D112lo12uFFJOJRuiOsf9NfYsG",
  "talkId": "wyyv2PswaGNyYGJnkBodXgd1daG6Rzxc",
  "answer": "{\"schemaVersion\":1,\"eventName\":\"ribbon.download.clicked\",\"source\":\"ribbon\"}"
}
```

字段约定：

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `frontEndIntent` | string | 固定值 `excelAi` |
| `clientSource` | string | 固定值 `Excel` |
| `questionType` | number | 固定值 `1` |
| `askId` | string | 每条事件随机生成，URL-safe |
| `talkId` | string | 每条事件随机生成，URL-safe |
| `answer` | string | 埋点事件 JSON 的字符串形式 |

埋点请求会复用插件 SSO 登录后的共享 cookie 容器。如果真实埋点接口和业务接口共用登录态，服务端可以直接按当前登录 cookie 鉴权。

当前实现位置：

- [src/OfficeAgent.Infrastructure/Analytics/InsertLogAnalyticsSink.cs](../src/OfficeAgent.Infrastructure/Analytics/InsertLogAnalyticsSink.cs)

## 3. `answer` 事件结构

`answer` 是一个 JSON 字符串。解析后的结构如下：

```json
{
  "schemaVersion": 1,
  "eventName": "ribbon.download.completed",
  "source": "ribbon",
  "occurredAtUtc": "2026-05-12T05:47:00.0000000Z",
  "properties": {
    "systemKey": "current-business-system",
    "projectId": "performance",
    "projectName": "绩效项目",
    "sheetName": "Sheet1",
    "workbookName": "demo.xlsx",
    "uiLocale": "zh",
    "operationScope": "partial",
    "rowCount": 10,
    "fieldCount": 4
  },
  "businessContext": {
    "module": "current-business-system",
    "endpoint": "/find"
  },
  "error": {
    "code": "operation_failed",
    "message": "Authentication required.",
    "exceptionType": "AuthenticationRequiredException"
  }
}
```

字段说明：

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `schemaVersion` | number | 当前固定为 `1` |
| `eventName` | string | 稳定事件名，格式为 `domain.action.phase` |
| `source` | string | 事件来源，例如 `ribbon`、`panel`、`connector` |
| `occurredAtUtc` | string | UTC 时间 |
| `properties` | object | 通用运营维度 |
| `businessContext` | object | 业务系统自定义上下文 |
| `error` | object | 失败事件的错误信息；成功事件通常没有该字段 |

## 4. `properties` 与 `businessContext` 怎么用

`properties` 放跨系统稳定字段，适合运营分析直接筛选和聚合：

- `systemKey`
- `projectId`
- `projectName`
- `sheetName`
- `workbookName`
- `uiLocale`
- `sessionId`
- `commandType`
- `operationScope`
- `rowCount`
- `fieldCount`
- `changeCount`
- `durationMs`

`businessContext` 放业务连接器自定义字段，适合真实系统内部解释事件：

- `module`
- `endpoint`
- `objectType`
- `activityType`
- `planType`
- `businessStatus`
- `fieldMappingCount`

规则：

- 能跨系统统一理解的字段放 `properties`
- 只有某个业务系统自己能理解的字段放 `businessContext`
- 不要把原始请求体、响应体、整行数据、单元格值放进 `businessContext`

## 5. 当前会产生哪些事件

常见事件前缀：

| 前缀 | 来源 | 示例 |
| --- | --- | --- |
| `panel.*` | 任务窗格 | `panel.composer.send.clicked`、`panel.settings.saved` |
| `ribbon.*` | Ribbon / Ribbon Sync | `ribbon.initialize.completed`、`ribbon.download.failed` |
| `connector.*` | Core 同步编排 | `connector.find.completed`、`connector.batch_save.failed` |
| `business.current.*` | 当前业务系统连接器 | `business.current.find.completed`、`business.current.schema.failed` |

具体事件名可参考：

- [docs/superpowers/specs/2026-05-12-office-agent-analytics-instrumentation-design.md](./superpowers/specs/2026-05-12-office-agent-analytics-instrumentation-design.md)
- [docs/modules/task-pane-current-behavior.md](./modules/task-pane-current-behavior.md)
- [docs/modules/ribbon-sync-current-behavior.md](./modules/ribbon-sync-current-behavior.md)

## 6. 敏感信息边界

当前实现的目标是只上报低敏结构化维度。

禁止上报：

- API key
- 把 cookie 内容写入 `answer`
- SSO token
- 用户 prompt 全文
- 单元格原始值
- 上传 / 下载业务接口 request body
- 上传 / 下载业务接口 response body
- 整行业务数据

允许上报：

- 是否配置了某类用户可见 URL，例如 `hasBaseUrl`、`hasBusinessBaseUrl`
- 用户输入长度，例如 `inputLength`
- 行数、字段数、变更数量、跳过数量
- `projectId`、`projectName`
- `sheetName`、`workbookName`
- 业务接口 endpoint 名称，例如 `/find`

注意：失败事件的 `error.message` 当前来自异常消息。真实连接器抛出的异常消息不要包含 token、请求体、响应体或单元格值。

## 7. 成功与失败处理

插件侧成功判断：

- HTTP `2xx` 视为成功
- 响应体内容不参与业务判断

失败处理：

- 非 `2xx`、网络异常、超时都会视为埋点失败
- 失败只写入本地诊断日志
- 不弹窗
- 不阻塞 Ribbon、任务窗格或同步业务流程

当前 HTTP 超时时间为 5 秒。埋点通过 fire-and-forget 方式异步发送。

## 8. 本地 mock 验证

启动 mock server：

```powershell
cd tests/mock-server
npm install
npm start
```

隐藏配置：

```text
AnalyticsUrl = http://localhost:3200/insertLog
```

即在 `%LocalAppData%\OfficeAgent\settings.json` 中保存 `"AnalyticsUrl": "http://localhost:3200/insertLog"`。

清空日志：

```powershell
Invoke-RestMethod -Method Delete -Uri http://localhost:3200/analytics/logs
```

触发 Excel 操作后查看日志：

```powershell
Invoke-RestMethod -Method Get -Uri http://localhost:3200/analytics/logs |
  ConvertTo-Json -Depth 20
```

如果先通过插件完成 SSO 登录，mock 日志里的单条记录会包含 `cookies` 字段，可用于确认埋点请求带上了登录 cookie。

手工 POST 验证：

```powershell
$body = @{
  frontEndIntent = "excelAi"
  clientSource = "Excel"
  questionType = 1
  askId = "ask-test"
  talkId = "talk-test"
  answer = '{"schemaVersion":1,"eventName":"panel.opened","source":"panel"}'
} | ConvertTo-Json -Compress

Invoke-RestMethod `
  -Method Post `
  -Uri http://localhost:3200/insertLog `
  -ContentType 'application/json' `
  -Body $body
```

## 9. 切换到真实系统

1. 确认真实系统完整埋点地址，例如：

   ```text
   https://analytics.internal.example/insertLog
   ```

2. 在 `%LocalAppData%\OfficeAgent\settings.json` 中保存：

   ```text
   AnalyticsUrl = https://analytics.internal.example/insertLog
   ```

3. 重启 Excel，确保 `ThisAddIn` 启动时按最新配置创建 `InsertLogAnalyticsSink`。

4. 触发一个低风险操作，例如打开任务窗格或点击 Ribbon `关于`。

5. 在真实系统侧确认收到：

   - `frontEndIntent = excelAi`
   - `clientSource = Excel`
   - `questionType = 1`
   - `answer` 可解析为 JSON
   - `answer.eventName` 为 `panel.*` 或 `ribbon.*`

6. 再触发 Ribbon Sync 初始化、下载、上传，确认 `projectId` 和 `projectName` 正常进入 `answer.properties`。

## 10. 真实系统验收清单

接口层：

- `POST /insertLog` 接受 `application/json`
- 固定字段校验与插件一致
- `answer` 按字符串接收，并能二次解析成 JSON
- HTTP 2xx 表示成功
- 非 2xx 响应体不包含敏感信息

插件配置：

- `AnalyticsUrl` 是完整埋点地址，包含 `/insertLog` 或真实系统要求的完整路径
- 空配置时不发送埋点
- 修改配置后重启 Excel 验证真实上报链路

数据质量：

- 至少能看到 `panel.*`、`ribbon.*`、`connector.*` 事件
- Ribbon Sync 事件包含 `projectId` 和 `projectName`
- Panel 发送事件只包含 `inputLength`，不包含 prompt 全文
- 上传 / 下载事件只包含计数和状态，不包含单元格值

安全：

- `answer` 中不出现 API key、cookie、token
- 不出现业务接口 request / response 原文
- 异常消息中不包含敏感数据

## 11. 常见问题

### 11.1 配了地址但没有上报

检查：

- `AnalyticsUrl` 是否为空
- 是否误填成只包含域名或基地址
- 真实埋点接口是否需要登录 cookie，插件是否已完成 SSO 登录
- Excel 是否在保存配置后重启
- `%LocalAppData%\OfficeAgent\logs\officeagent.log` 是否有 `analytics track.failed`

### 11.2 真实系统收到的 `answer` 是字符串

这是预期行为。内网接口合同要求 `answer` 是字符串，所以插件会把内部事件 JSON 序列化后放入 `answer` 字段。

真实系统如果要分析事件，需要对 `answer` 再做一次 JSON parse。

### 11.3 `askId` 和 `talkId` 是否等于会话 ID

当前不是。它们是每条上报随机生成的 URL-safe ID。

如果后续真实系统要求 `talkId` 绑定任务窗格会话或 Excel 会话，需要再扩展 `InsertLogAnalyticsSink` 的 ID 生成策略。

### 11.4 能否在业务连接器里加自定义字段

可以。业务连接器注入 `IAnalyticsService` 后，可以调用：

```csharp
analyticsService.Track(
    "business.real.find.completed",
    "connector",
    properties,
    businessContext);
```

稳定字段放 `properties`，业务私有字段放 `businessContext`。

### 11.5 埋点失败会不会影响用户操作

不会。埋点发送失败只写本地诊断日志，不改变用户当前操作结果。
