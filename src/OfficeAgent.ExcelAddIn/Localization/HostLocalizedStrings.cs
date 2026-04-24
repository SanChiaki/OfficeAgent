using System;

namespace OfficeAgent.ExcelAddIn.Localization
{
    public sealed class HostLocalizedStrings
    {
        private HostLocalizedStrings(string locale)
        {
            Locale = string.Equals(locale, "zh", StringComparison.OrdinalIgnoreCase) ? "zh" : "en";
        }

        public string Locale { get; }

        public string HostWindowTitle => "ISDP";

        public string RibbonTabLabel => "ISDP";

        public string RibbonAgentGroupLabel => "ISDP AI";

        public string RibbonAgentButtonLabel => "ISDP AI";

        public string RibbonProjectGroupLabel => Locale == "zh" ? "项目" : "Project";

        public string ProjectDropDownPlaceholderText => Locale == "zh" ? "先选择项目" : "Select project";

        public string ProjectDropDownLoginRequiredText => Locale == "zh" ? "请先登录" : "Sign in first";

        public string ProjectDropDownNoAvailableProjectsText => Locale == "zh" ? "无可用项目" : "No projects available";

        public string ProjectDropDownLoadFailedText => Locale == "zh" ? "项目加载失败" : "Failed to load projects";

        public string RibbonInitializeSheetButtonLabel => Locale == "zh" ? "初始化当前表" : "Initialize sheet";

        public string RibbonTemplateGroupLabel => Locale == "zh" ? "模板" : "Template";

        public string RibbonApplyTemplateButtonLabel => Locale == "zh" ? "应用模板" : "Apply template";

        public string RibbonSaveTemplateButtonLabel => Locale == "zh" ? "保存模板" : "Save template";

        public string RibbonSaveAsTemplateButtonLabel => Locale == "zh" ? "另存模板" : "Save as template";

        public string RibbonDownloadGroupLabel => Locale == "zh" ? "下载" : "Download";

        public string RibbonPartialDownloadButtonLabel => Locale == "zh" ? "部分下载" : "Partial download";

        public string RibbonFullDownloadButtonLabel => Locale == "zh" ? "全量下载" : "Full download";

        public string RibbonUploadGroupLabel => Locale == "zh" ? "上传" : "Upload";

        public string RibbonPartialUploadButtonLabel => Locale == "zh" ? "部分上传" : "Partial upload";

        public string RibbonFullUploadButtonLabel => Locale == "zh" ? "全量上传" : "Full upload";

        public string RibbonAccountGroupLabel => Locale == "zh" ? "账号" : "Account";

        public string RibbonLoginButtonLabel => Locale == "zh" ? "登录" : "Login";

        public string RibbonLoginInProgressButtonLabel => Locale == "zh" ? "登录中..." : "Signing in...";

        public string ConfigureSsoUrlFirstMessage => Locale == "zh"
            ? "请先在设置中配置 SSO 地址。"
            : "Configure the SSO URL in Settings first.";

        public string AuthenticationRequiredDefaultMessage => Locale == "zh"
            ? "当前未登录，请先登录"
            : "You're not signed in. Sign in first.";

        public string AuthenticationRequiredLoginButtonText => Locale == "zh" ? "点我登录" : "Sign in";

        public string CloseButtonText => Locale == "zh" ? "关闭" : "Close";

        public string OkButtonText => "OK";

        public string CancelButtonText => Locale == "zh" ? "取消" : "Cancel";

        public string SsoLoginPopupTitle => Locale == "zh" ? "ISDP - 登录" : "ISDP - Sign in";

        public string SsoLoginConfirmedButtonText => Locale == "zh" ? "已登录" : "I've signed in";

        public string ProjectListEmptyWarningMessage => Locale == "zh"
            ? "项目列表加载完成，但未获取到任何可用项目。\r\n请检查登录状态或项目接口返回。"
            : "The project list loaded, but no projects are available.\r\nCheck your sign-in state or project API response.";

        public string ProjectListLoadFailedMessage(string details)
        {
            return Locale == "zh"
                ? $"项目列表加载失败。\r\n{details}"
                : $"Failed to load projects.\r\n{details}";
        }

        public string ConfirmOperationPrompt(string operationName)
        {
            return Locale == "zh"
                ? $"确认要执行{operationName}吗？"
                : $"Run {operationName}?";
        }

        public string ProjectLine(string projectName)
        {
            return Locale == "zh"
                ? $"项目：{projectName}"
                : $"Project: {projectName}";
        }

        public string RowCountLine(int rowCount)
        {
            return Locale == "zh"
                ? $"记录数：{rowCount}"
                : $"Rows: {rowCount}";
        }

        public string FieldCountLine(int fieldCount)
        {
            return Locale == "zh"
                ? $"字段数：{fieldCount}"
                : $"Fields: {fieldCount}";
        }

        public string OverwriteDirtyCellsLine(int dirtyCount)
        {
            return Locale == "zh"
                ? $"将覆盖 {dirtyCount} 个未上传改单元格。"
                : $"This will overwrite {dirtyCount} unsaved edited cells.";
        }

        public string ProjectLayoutDialogTitle => Locale == "zh" ? "配置当前表布局" : "Configure sheet layout";

        public string ProjectLayoutInstructionText => Locale == "zh"
            ? "下面三个值会写入当前工作表的同步配置（SheetBindings），请确认后保存。"
            : "The three values below will be saved to the current worksheet sync settings (SheetBindings). Review them before saving.";

        public string ProjectLayoutCurrentBindingText(string projectId, string projectName)
        {
            return Locale == "zh"
                ? $"当前绑定：{projectId} | {projectName}"
                : $"Current binding: {projectId} | {projectName}";
        }

        public string ProjectLayoutPositiveIntegerError(string fieldName)
        {
            return Locale == "zh"
                ? $"{fieldName} 必须是正整数。"
                : $"{fieldName} must be a positive integer.";
        }

        public string ProjectLayoutDataStartValidationError => Locale == "zh"
            ? "DataStartRow 必须大于或等于 HeaderStartRow + HeaderRowCount。"
            : "DataStartRow must be greater than or equal to HeaderStartRow + HeaderRowCount.";

        public string TaskPaneRuntimeMissingMessage => Locale == "zh"
            ? "需要 WebView2 Runtime 才能显示 ISDP。"
            : "WebView2 Runtime is required to render ISDP.";

        public string TaskPaneInitializationFailedMessage => Locale == "zh"
            ? "ISDP 无法初始化任务窗格。请检查本地日志后重新打开 Excel。"
            : "ISDP could not initialize the task pane. Check the local log and reopen Excel.";

        public string BridgeUnexpectedErrorMessage => Locale == "zh"
            ? "ISDP 遇到了未预期的错误。请检查本地日志后重试。"
            : "ISDP hit an unexpected error. Check the local log and try again.";

        public string BridgeMalformedJsonMessage => Locale == "zh"
            ? "Web 消息 payload 不是有效的 JSON。"
            : "The web message payload was not valid JSON.";

        public string BridgeMalformedRequestMessage => Locale == "zh"
            ? "Web 消息必须同时包含 type 和 requestId。"
            : "Web messages must include both type and requestId.";

        public string BridgeUnknownMessageTypeMessage(string messageType)
        {
            return Locale == "zh"
                ? $"消息类型“{messageType}”不被允许。"
                : $"Message type '{messageType}' is not allowed.";
        }

        public string BridgePayloadNotAcceptedMessage(string bridgeType)
        {
            return Locale == "zh"
                ? $"{bridgeType} 不接受 payload。"
                : $"{bridgeType} does not accept a payload.";
        }

        public string BridgePayloadRequiredMessage(string bridgeType, string payloadDescription)
        {
            return Locale == "zh"
                ? $"{bridgeType} 需要{payloadDescription}。"
                : $"{bridgeType} requires {payloadDescription}.";
        }

        public string BridgeValidPayloadRequiredMessage(string bridgeType, string payloadDescription)
        {
            return Locale == "zh"
                ? $"{bridgeType} 需要有效的{payloadDescription}。"
                : $"{bridgeType} requires a valid {payloadDescription}.";
        }

        public string BridgeLoginMustBeAsyncMessage => Locale == "zh"
            ? "bridge.login 必须走异步路由。"
            : "bridge.login must be routed asynchronously.";

        public string BridgeMissingSsoUrlMessage => Locale == "zh"
            ? "请先配置 SSO URL。"
            : "Configure the SSO URL first.";

        public string BridgeLoginCanceledMessage => Locale == "zh"
            ? "用户取消了登录。"
            : "The sign-in flow was canceled.";

        public string BridgeBusyMessage => Locale == "zh"
            ? "已有请求正在处理中，请稍候。"
            : "Another request is already in progress. Please wait.";

        public string BridgeAgentRequestTimedOutMessage => Locale == "zh"
            ? "Agent 请求已超时。"
            : "Agent request timed out.";

        public string BridgeAgentExecutionFailedMessage => Locale == "zh"
            ? "Agent 执行失败。"
            : "Agent execution failed.";

        public string BootstrapperFallbackHtml => Locale == "zh"
            ? @"<!doctype html>
<html lang=""zh-CN"">
  <head>
    <meta charset=""utf-8"" />
    <title>ISDP</title>
    <style>
      body { font-family: Segoe UI, sans-serif; padding: 24px; color: #1f2937; }
      code { background: #f3f4f6; padding: 2px 6px; border-radius: 4px; }
    </style>
  </head>
  <body>
    <h1>ISDP</h1>
    <p>未找到前端资源。</p>
    <p>请先构建 <code>src/OfficeAgent.Frontend</code>，然后重新打开任务窗格。</p>
  </body>
</html>"
            : @"<!doctype html>
<html lang=""en"">
  <head>
    <meta charset=""utf-8"" />
    <title>ISDP</title>
    <style>
      body { font-family: Segoe UI, sans-serif; padding: 24px; color: #1f2937; }
      code { background: #f3f4f6; padding: 2px 6px; border-radius: 4px; }
    </style>
  </head>
  <body>
    <h1>ISDP</h1>
    <p>Frontend assets were not found.</p>
    <p>Build <code>src/OfficeAgent.Frontend</code> and reopen the task pane.</p>
  </body>
</html>";

        public static HostLocalizedStrings ForLocale(string locale)
        {
            return new HostLocalizedStrings(locale);
        }
    }
}
