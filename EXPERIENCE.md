# EXPERIENCE.md

VSTO + WebView2 + Excel 开发中踩过的坑和经验总结。

## 1. VSTO 中 async/await 的 SynchronizationContext 陷阱

**现象：** `ConfigureAwait(true)` 在 VSTO 环境中静默失效——await 之后的代码跑在线程池线程上，而不是 UI 线程。调用 `CoreWebView2.PostWebMessageAsJson` 时抛出 "CoreWebView2 can only be accessed from the UI thread"。

**根因：** VSTO 不调用 `Application.Run()`，因此 `WindowsFormsSynchronizationContext` 不会自动安装到 Excel 的 STA 线程上。`SynchronizationContext.Current` 为 `null`，导致 `ConfigureAwait(true)` 无上下文可捕获，续体直接在完成 await 的线程池线程上执行。

**修复：** 在 `WebViewBootstrapper.InitializeAsync()` 的第一个 `await` 之前，显式安装 `WindowsFormsSynchronizationContext`：

```csharp
if (!(SynchronizationContext.Current is WindowsFormsSynchronizationContext))
{
    SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
}
```

**教训：** 在 VSTO/Office Add-in 环境中，永远不要假设 `SynchronizationContext` 已经就位。安装一次后，整条 async 链的 `ConfigureAwait(true)` 都能正常工作。

## 2. async void 事件处理器中的异常会杀掉宿主进程

**现象：** Excel 闪退，没有任何错误对话框。

**根因：** `CoreWebView2.PostWebMessageAsJson` 在 WebView2 渲染进程崩溃或状态异常时会抛异常。在 `async void` 事件处理器中，如果 `catch` 块里调用的 `PostError` 也抛了，异常逃逸出 `async void`，.NET Framework 会直接终止进程。

**修复：** 所有 `PostWebMessageAsJson` 调用都走 `TryPostWebMessage` 方法，内部 try/catch + null-conditional：

```csharp
private void TryPostWebMessage(string json)
{
    try
    {
        if (webView.InvokeRequired)
            webView.Invoke((Action)(() => webView.CoreWebView2?.PostWebMessageAsJson(json)));
        else
            webView.CoreWebView2?.PostWebMessageAsJson(json);
    }
    catch (Exception error)
    {
        OfficeAgentLog.Warn("webview", "post.failed", $"PostWebMessage failed: {error.Message}");
    }
}
```

**教训：**
- `async void` 中的每一条路径（包括 catch 块）都必须防住二次异常
- COM 事件回调（如 `SheetSelectionChange`）中调用 WebView2 也必须 try/catch，否则异常传回 Excel 的 COM dispatcher 会闪退
- 用 `webView.InvokeRequired` 作为最后一道防线，即使 `SynchronizationContext` 出问题也能正确回到 UI 线程

## 3. sync-over-async 阻塞 Excel UI 线程

**现象：** Agent "思考中"时，Excel 整个界面冻结，无法编辑单元格。

**根因：** `WebMessageReceived` 事件在 UI/STA 线程上触发，`HttpClient.SendAsync().GetAwaiter().GetResult()` 同步阻塞该线程 30-120 秒。STA 线程被阻塞后无法处理 Windows 消息，Excel 界面完全无响应。

**修复：** 整条调用链改为 async：
- `WebViewBootstrapper.CoreWebView2_WebMessageReceived` → `async void`
- `WebMessageRouter.RouteAsync` → `Task<string>`
- `AgentOrchestrator.ExecuteAsync` → `Task<AgentCommandResult>`
- `LlmPlannerClient.CompleteAsync` → `Task<string>`
- HTTP 层用 `ConfigureAwait(false)`，Orchestrator 层用 `ConfigureAwait(true)`

**教训：** 在 STA 线程上（VSTO、WinForms、WPF）永远不要对 HTTP 或 I/O 操作做 `.Result` / `.GetAwaiter().GetResult()`。要么全程 async，要么用 `Task.Run` 配合显式 `SynchronizationContext.Post` 回到 UI 线程（后者更复杂，不推荐）。

## 4. HttpClient 超时与 TaskCanceledException

**现象：** LLM 调用 30 秒后抛出 `TaskCanceledException`。

**根因：** `HttpClient.Timeout` 默认 100 秒，但代码中显式设了 30 秒。某些 LLM 请求（尤其是多步 plan）需要更长时间。

**修复：** 将 `HttpClient.Timeout` 增加到 120 秒。

**教训：** 在诊断超时问题时，先看 `HttpClient.Timeout` 的值，再考虑网络问题。`TaskCanceledException` 不一定是用户取消，更可能是超时。

## 5. MSI 版本号写死导致无法覆盖安装

**现象：** 新版 MSI 安装时提示已安装，无法覆盖。

**根因：** WiX `Package` 元素的 `Version="1.0.0.0"` 写死。Windows Installer 比较版本号，相同版本不触发 MajorUpgrade。

**修复：** 版本号从 git commit count 自动生成：
- `Product.wxs`: `Version="$(var.ProductVersion)"`
- `build.ps1`: `$productVersion = "1.0.$(git rev-list --count HEAD)"`

**注意：** Windows Installer 版本比较只看前三段（`major.minor.build`），第四段 `revision` 被忽略。所以必须把递增的数字放在第三段（build），不能用 `1.0.0.N` 格式。

## 6. WebView2 渲染进程崩溃的诊断

**现象：** Excel 操作一段时间后突然闪退，日志中没有 bridge 错误。

**排查方法：**
1. 检查 `%LocalAppData%\OfficeAgent\logs\officeagent.log`
2. 如果最后一条日志是 `bridge.runAgent` 但没有对应的完成日志，说明 crash 发生在 await 期间
3. 检查日志中是否有 `webview.post.failed` 或 `webview.process.failed` 警告
4. 注册 `CoreWebView2.ProcessFailed` 事件可以在日志中捕获渲染进程崩溃

## 7. VSTO 构建需要 Visual Studio MSBuild

**现象：** `dotnet build` 对 VSTO 项目报错（找不到 SmartTagCollection、Office Interop 等）。

**原因：** VSTO 项目依赖 Visual Studio 的专用 targets 和引用程序集，`dotnet build` 不具备这些。

**正确做法：**
- 构建用 `MSBuild.exe`（Visual Studio 自带）：`"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"`
- 测试用 `dotnet test`（可以正常工作）
- MSI 构建脚本 `build.ps1` 内部已调用正确的 MSBuild 路径

## 8. Excel COM Interop 的线程安全

**规则：**
- Excel COM 对象只能在创建它们的 STA 线程上访问
- `Application.SheetSelectionChange` 事件在 STA 线程上触发
- `WebView2.WebMessageReceived` 事件在创建 WebView2 的 STA 线程上触发
- 任何 `await` 之后的 COM 调用，都必须确保续体回到 STA 线程（`ConfigureAwait(true)` + `WindowsFormsSynchronizationContext`）

**常见错误模式：**
```csharp
// ❌ 错误：Task.Run 中的 COM 调用会抛 InvalidComObjectException
Task.Run(() => excelContextService.GetCurrentSelectionContext());

// ✅ 正确：COM 调用在 await 之前（同步），HTTP 调用在 await 之后
var context = excelContextService.GetCurrentSelectionContext(); // 在 UI 线程
var response = await httpClient.SendAsync(request);              // 释放 UI 线程
// ConfigureAwait(true) 确保 return 后回到 UI 线程
```

## 9. 前端 bridge 测试策略

前端测试中所有 bridge 调用都 mock 了 `window.chrome.webview`。测试验证的是：
- 消息格式（`type`、`requestId`、`payload`）
- UI 状态变化（loading 显示/隐藏、确认卡片出现/消失）
- 用户交互（发送、确认、取消）

**注意：** 前端 UI 文案已翻译为中文，测试断言必须匹配中文字符串。修改 UI 文案时务必同步更新测试。

## 10. PowerShell 脚本中的构建命令

`build.ps1` 中调用外部命令使用 `Invoke-NativeCommand` 封装，检查 `$LASTEXITCODE`。注意：
- PowerShell 中 `git` 命令返回的输出可能有尾部空白，需要 `.Trim()`
- WiX 变量通过 `-d Name=Value` 传入，在 `.wxs` 中用 `$(var.Name)` 引用
- 路径分隔符在 PowerShell 中用 `\` 或 `\\`（在双引号字符串中）

## 11. VSTO ClickOnce 清单签名在 CI 中的处理

**现象：** GitHub Actions 构建 VSTO 项目时报错 "Cannot build because the ClickOnce manifest signing option is not selected."

**根因：** VSTO 项目（`.csproj`）默认 `SignManifests=true` 并绑定了开发机的证书指纹。CI 环境没有该证书，且 `/p:SignManifests=false` 也不能用——VSTO targets 强制要求签名。

**修复：** 在 CI workflow 中生成临时自签名证书：
```powershell
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=OfficeAgent CI Test" -CertStoreLocation "Cert:\CurrentUser\My"
echo "SIGNING_THUMBPRINT=$($cert.Thumbprint)" >> $env:GITHUB_ENV
```
然后通过 `/p:ManifestCertificateThumbprint=$SIGNING_THUMBPRINT` 传给 MSBuild。构建结束后清理证书。

**教训：** VSTO 清单签名不可跳过，CI 必须提供有效证书。本地开发机上的证书指纹在 CI 上不可用。

## 12. WiX v4 需要通过 .NET 工具清单安装

**现象：** `build.ps1` 中 `dotnet tool restore` 在全新 clone 上失败。

**根因：** 项目缺少 `.dotnet-tools.json` 工具清单文件。WiX v4 作为 .NET 全局/本地工具分发（`wix` NuGet 包 v4.0.5），必须有清单文件声明依赖。

**修复：** 在 worktree 根目录添加 `.dotnet-tools.json`：
```json
{
  "version": 1,
  "isRoot": true,
  "tools": {
    "wix": { "version": "4.0.5", "commands": ["wix"] }
  }
}
```

## 13. Agent 对话历史注入点

**背景：** LLM planner 原本只收到当前用户消息（无状态），无法理解追问和指代。

**设计决策：** 前端发送历史而非后端加载。理由：
- 前端 React state 已持有完整消息列表（`sessionThreads`），是消息的真实来源
- 避免在 orchestrator 中耦合文件 I/O（`FileSessionStore`）
- 保持 `LlmPlannerClient` 无状态、可测试

**注入路径：** `AgentCommandEnvelope.ConversationHistory` → `PlannerRequest.ConversationHistory` → `LlmPlannerClient.BuildChatMessages`。消息数组结构：`[system, ...history, current_user]`。

**Token 控制：** 前端裁剪最近 10 轮（20 条消息），只取 `user`/`assistant` 角色的纯文本内容，不取完整结构化 JSON。

**注意事项：**
- `ConversationHistory` 默认空数组，向后兼容——所有现有测试无需修改
- assistant 消息只存 `AssistantMessage` 文本，不存完整 `PlannerResponse`（避免 token 浪费）
- `build.ps1` 中 MSBuild 路径硬编码为 VS 2022 Community 版本，CI（Enterprise 版）需要 `microsoft/setup-msbuild` action 或 `vswhere` 动态查找

## 14. 会话管理的保存策略

**设计决策：** 前端发起 `bridge.saveSessions`，后端只负责接收和写入文件。

**理由：**
- 前端 React state 是会话消息的真实来源（`sessionThreads`），包含了 system 消息、table 附件等后端不需要的信息
- 避免在 orchestrator 中耦合文件 I/O
- 所有会话管理操作（创建、重命名、删除、切换）在前端完成，通过 `saveCurrentSessions()` 持久化

**保存时机：**
- `sessionThreads` 变化时 1 秒防抖自动保存（useEffect）
- 切换会话前立即保存
- 创建新会话前立即保存
- 删除会话后立即保存

**注意事项：**
- `saveCurrentSessions()` 是 async 函数，用 `void` 调用（fire-and-forget），错误静默吞掉
- `threadToChatMessages` 只取 `user`/`assistant` 角色，过滤掉 `system` 角色和 `table` 字段

## 15. VSTO Ribbon 按钮图标

**问题：** Ribbon 按钮默认无图标，需要自定义 logo。

**方案：** 将 logo 图片以 base64 嵌入 `Properties/Resources.resx`，在 `Ribbon_Load` 中通过 `Properties.Resources.Logo` 获取 `System.Drawing.Bitmap` 赋值给按钮的 `Image` 属性。

**不推荐的方案：**
- 从文件系统加载（`Image.FromFile`）：VSTO 部署路径不确定，文件可能找不到
- 嵌入为 `EmbeddedResource` 再 `Assembly.GetManifestResourceStream`：多一步转换，不如 resx 直接

**注意：** Ribbon 大按钮图标推荐 32x32 PNG；图片需要有透明背景（PNG alpha）。
