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

        public string HostWindowTitle => "xISDP";

        public string RibbonTabLabel => "xISDP";

        public string RibbonAgentGroupLabel => "xISDP AI";

        public string RibbonAgentButtonLabel => string.Empty;

        public string RibbonProjectGroupLabel => Locale == "zh" ? "项目" : "Project";

        public string ProjectDropDownPlaceholderText => Locale == "zh" ? "先选择项目" : "Select project";

        public string ProjectDropDownLoginRequiredText => Locale == "zh" ? "请先登录" : "Sign in first";

        public string ProjectDropDownNoAvailableProjectsText => Locale == "zh" ? "无可用项目" : "No projects available";

        public string ProjectDropDownLoadFailedText => Locale == "zh" ? "项目加载失败" : "Failed to load projects";

        public string ProjectSelectionRequiredMessage => Locale == "zh" ? "请先选择项目。" : "Select a project first.";

        public bool IsStickyProjectStatus(string text)
        {
            return string.Equals(text, ProjectDropDownLoginRequiredText, StringComparison.Ordinal) ||
                   string.Equals(text, ProjectDropDownNoAvailableProjectsText, StringComparison.Ordinal) ||
                   string.Equals(text, ProjectDropDownLoadFailedText, StringComparison.Ordinal);
        }

        public string RibbonInitializeSheetButtonLabel => Locale == "zh" ? "初始化当前表" : "Initialize sheet";

        public string RibbonAiMapColumnsButtonLabel => Locale == "zh" ? "AI映射列" : "AI map columns";

        public string RibbonTemplateGroupLabel => Locale == "zh" ? "配置" : "Setting";

        public string RibbonApplyTemplateButtonLabel => Locale == "zh" ? "应用配置" : "Apply Setting";

        public string RibbonSaveTemplateButtonLabel => Locale == "zh" ? "保存配置" : "Save Setting";

        public string RibbonSaveAsTemplateButtonLabel => Locale == "zh" ? "另存配置" : "Save as Setting";

        public string InitializeCurrentSheetCompletedMessage => Locale == "zh"
            ? "初始化当前表完成。"
            : "Initialize sheet completed.";

        public string AiColumnMappingPreviewDialogTitle => Locale == "zh" ? "确认 AI 映射列" : "Confirm AI column mapping";

        public string AiColumnMappingProgressDialogTitle => Locale == "zh" ? "AI 映射列处理中" : "AI column mapping";

        public string AiColumnMappingProgressMessage => Locale == "zh"
            ? "AI 正在分析当前表头并生成列映射，请等待处理完成。"
            : "AI is analyzing the current headers and generating column mappings. Wait for it to finish.";

        public string AiColumnMappingAbortButtonText => Locale == "zh" ? "中止" : "Abort";

        public string AiColumnMappingPreviewInstructionText => Locale == "zh"
            ? "请选择需要写入的 AI 推荐列映射。确认后仅会更新 xISDP_Setting.SheetFieldMappings 中的 Excel L1 / Excel L2。"
            : "Select the AI column mappings to write. Confirming updates only Excel L1 / Excel L2 in xISDP_Setting.SheetFieldMappings.";

        public string AiColumnMappingApplyColumnHeader => Locale == "zh" ? "是否修改" : "Apply";

        public string AiColumnMappingExcelColumnHeader => Locale == "zh" ? "列号" : "Column";

        public string AiColumnMappingActualHeaderColumnHeader => Locale == "zh" ? "当前表头（L1/L2）" : "Current header (L1/L2)";

        public string AiColumnMappingMatchedHeaderColumnHeader => Locale == "zh" ? "匹配表头（L1/L2）" : "Matched header (L1/L2)";

        public string AiColumnMappingNoAcceptedMappingsMessage => Locale == "zh"
            ? "AI 映射列没有可应用的推荐。"
            : "AI column mapping found no accepted mappings.";

        public string AiColumnMappingCompletedMessage(int appliedCount, int skippedCount)
        {
            return Locale == "zh"
                ? $"AI 映射列完成。\r\n已应用：{appliedCount}\r\n已跳过：{skippedCount}"
                : $"AI column mapping completed.\r\nApplied: {appliedCount}\r\nSkipped: {skippedCount}";
        }

        public string RibbonDataSyncGroupLabel => Locale == "zh" ? "数据同步" : "Data sync";

        public string RibbonPartialDownloadButtonLabel => Locale == "zh" ? "下载" : "Download";

        public string RibbonFullDownloadButtonLabel => Locale == "zh" ? "全量下载" : "Full download";

        public string RibbonPartialUploadButtonLabel => Locale == "zh" ? "上传" : "Upload";

        public string RibbonFullUploadButtonLabel => Locale == "zh" ? "全量上传" : "Full upload";

        public string RibbonHelpGroupLabel => Locale == "zh" ? "帮助" : "Help";

        public string RibbonDocumentationButtonLabel => Locale == "zh" ? "文档" : "Documentation";

        public string RibbonAboutButtonLabel => Locale == "zh" ? "关于" : "About";

        public string RibbonAccountGroupLabel => Locale == "zh" ? "账号" : "Account";

        public string RibbonLoginButtonLabel => Locale == "zh" ? "登录" : "Login";

        public string RibbonLoginInProgressButtonLabel => Locale == "zh" ? "登录中..." : "Signing in...";

        public string ConfigureSsoUrlFirstMessage => Locale == "zh"
            ? "请先在设置中配置 SSO 地址。"
            : "Configure the SSO URL in Settings first.";

        public string DocumentationOpenFailedMessage(string details)
        {
            return Locale == "zh"
                ? $"无法打开文档页面。\r\n{details}"
                : $"Could not open the documentation page.\r\n{details}";
        }

        public string RibbonAboutDialogTitle => Locale == "zh" ? "关于 xISDP" : "About xISDP";

        public string UnknownText => Locale == "zh" ? "未知" : "Unknown";

        public string AboutMessage(string appVersion, string assemblyVersion, string buildConfiguration, string buildTime)
        {
            return Locale == "zh"
                ? "OfficeAgent Excel Add-in\r\n" +
                    "版本号: " + appVersion + "\r\n" +
                    "程序集版本: " + assemblyVersion + "\r\n" +
                    "构建配置: " + buildConfiguration + "\r\n" +
                    "构建时间: " + buildTime
                : "OfficeAgent Excel Add-in\r\n" +
                    "Version: " + appVersion + "\r\n" +
                    "Assembly version: " + assemblyVersion + "\r\n" +
                    "Build configuration: " + buildConfiguration + "\r\n" +
                    "Build time: " + buildTime;
        }

        public string AuthenticationRequiredDefaultMessage => Locale == "zh"
            ? "当前未登录，请先登录"
            : "You're not signed in. Sign in first.";

        public string AuthenticationRequiredLoginButtonText => Locale == "zh" ? "点我登录" : "Sign in";

        public string CloseButtonText => Locale == "zh" ? "关闭" : "Close";

        public string OkButtonText => "OK";

        public string CancelButtonText => Locale == "zh" ? "取消" : "Cancel";

        public string SsoLoginPopupTitle => Locale == "zh" ? "xISDP - 登录" : "xISDP - Sign in";

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

        public string SubmittedCellCountLine(int cellCount)
        {
            return Locale == "zh"
                ? $"提交单元格数：{cellCount}"
                : $"Submitted cells: {cellCount}";
        }

        public string SkippedCellCountLine(int cellCount)
        {
            return Locale == "zh"
                ? $"跳过单元格数：{cellCount}"
                : $"Skipped cells: {cellCount}";
        }

        public string OverwriteDirtyCellsLine(int dirtyCount)
        {
            return Locale == "zh"
                ? $"将覆盖 {dirtyCount} 个未上传改单元格。"
                : $"This will overwrite {dirtyCount} unsaved edited cells.";
        }

        public string ProjectLayoutDialogTitle => Locale == "zh" ? "配置当前表布局" : "Configure sheet layout";

        public string ProjectLayoutInstructionText => Locale == "zh"
            ? "下面三个值会写入当前工作表与ISDP实施计划的映射配置表xISDP_Setting中，请确认后保存。"
            : "The three values below will be written to xISDP_Setting, the mapping configuration table for the current worksheet and the ISDP implementation plan. Confirm them before saving.";

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

        public string DefaultTemplateDisplayName => Locale == "zh" ? "未绑定模板" : "No template linked";

        public string TemplatePickerDialogTitle => Locale == "zh" ? "应用模板" : "Apply template";

        public string TemplatePickerCurrentProjectText(string projectDisplayName)
        {
            return Locale == "zh"
                ? $"当前项目：{projectDisplayName}"
                : $"Current project: {projectDisplayName}";
        }

        public string TemplatePickerInstructionText => Locale == "zh"
            ? "请选择要应用到当前表的本机模板。"
            : "Select a local template to apply to the current sheet.";

        public string TemplatePickerSelectionRequiredMessage => Locale == "zh"
            ? "请选择一个模板。"
            : "Select a template.";

        public string TemplateNoAvailableMessage => Locale == "zh"
            ? "当前项目没有可用模板。"
            : "No templates are available for the current project.";

        public string TemplateNotFoundMessage => Locale == "zh"
            ? "未找到所选模板。"
            : "The selected template was not found.";

        public string ApplyTemplateCompletedMessage(string templateName)
        {
            return Locale == "zh"
                ? $"应用模板完成。\r\n模板：{templateName}"
                : $"Apply template completed.\r\nTemplate: {templateName}";
        }

        public string TemplateNoSavableMessage => Locale == "zh"
            ? "当前表没有可保存的模板。"
            : "The current sheet has no template to save.";

        public string SaveTemplateCompletedMessage(string templateName)
        {
            return Locale == "zh"
                ? $"保存模板完成。\r\n模板：{templateName}"
                : $"Save template completed.\r\nTemplate: {templateName}";
        }

        public string OverwriteTemplateCompletedMessage(string templateName)
        {
            return Locale == "zh"
                ? $"覆盖模板完成。\r\n模板：{templateName}"
                : $"Overwrite template completed.\r\nTemplate: {templateName}";
        }

        public string SuggestedNewTemplateName => Locale == "zh" ? "新模板" : "New template";

        public string FormatSuggestedTemplateCopyName(string templateName)
        {
            if (string.IsNullOrWhiteSpace(templateName))
            {
                return SuggestedNewTemplateName;
            }

            return Locale == "zh"
                ? templateName + "-副本"
                : templateName + "-copy";
        }

        public string SaveAsTemplateCompletedMessage(string templateName)
        {
            return Locale == "zh"
                ? $"另存模板完成。\r\n模板：{templateName}"
                : $"Save as template completed.\r\nTemplate: {templateName}";
        }

        public string TemplateNameDialogTitle => Locale == "zh" ? "另存模板" : "Save as template";

        public string TemplateNameDialogPrompt => Locale == "zh"
            ? "请输入新模板名称。保存后，当前表会绑定到新模板。"
            : "Enter a new template name. After saving, the current sheet will be linked to the new template.";

        public string TemplateNameRequiredMessage => Locale == "zh"
            ? "模板名称不能为空。"
            : "Template name cannot be empty.";

        public string TemplateOverwriteConfirmTitle => Locale == "zh" ? "覆盖模板" : "Overwrite template";

        public string TemplateOverwriteConfirmMessage(string templateName)
        {
            return Locale == "zh"
                ? $"当前表存在未保存的模板改动，确认用模板“{templateName}”覆盖吗？"
                : $"The current sheet has unsaved template changes. Overwrite it with template \"{templateName}\"?";
        }

        public string TemplateOverwriteButtonText => Locale == "zh" ? "覆盖" : "Overwrite";

        public string TemplateRevisionConflictTitle => Locale == "zh" ? "模板版本冲突" : "Template revision conflict";

        public string TemplateRevisionConflictMessage(string templateName, int sheetRevision, int storedRevision)
        {
            return Locale == "zh"
                ? $"模板“{templateName}”已从版本 {sheetRevision} 更新到版本 {storedRevision}。\r\n请选择后续操作。"
                : $"Template \"{templateName}\" changed from revision {sheetRevision} to {storedRevision}.\r\nChoose what to do next.";
        }

        public string TemplateSaveAsNewButtonText => Locale == "zh" ? "另存为新模板" : "Save as new template";

        public string TemplateOverwriteOriginalButtonText => Locale == "zh" ? "覆盖原模板" : "Overwrite original template";

        public string LocalizeSyncOperationName(string operationName)
        {
            switch ((operationName ?? string.Empty).Trim())
            {
                case "全量下载":
                    return RibbonFullDownloadButtonLabel;
                case "部分下载":
                    return RibbonPartialDownloadButtonLabel;
                case "全量上传":
                    return RibbonFullUploadButtonLabel;
                case "部分上传":
                    return RibbonPartialUploadButtonLabel;
                default:
                    return operationName ?? string.Empty;
            }
        }

        public string FormatDownloadCompletedMessage(string operationName, int rowCount, int fieldCount)
        {
            var localizedOperationName = LocalizeSyncOperationName(operationName);
            return Locale == "zh"
                ? $"{localizedOperationName}完成。\r\n{RowCountLine(rowCount)}\r\n{FieldCountLine(fieldCount)}"
                : $"{localizedOperationName} completed.\r\n{RowCountLine(rowCount)}\r\n{FieldCountLine(fieldCount)}";
        }

        public string FormatDownloadNoMatchingRowsMessage(string operationName)
        {
            return Locale == "zh"
                ? "查询结果为空，请确认列名是否正确匹配。"
                : "The query result is empty. Check whether the column names are mapped correctly.";
        }

        public string FormatUploadNoChangesMessage(string operationName)
        {
            var localizedOperationName = LocalizeSyncOperationName(operationName);
            return Locale == "zh"
                ? $"{localizedOperationName}没有可提交的单元格。"
                : $"{localizedOperationName} has no cells to submit.";
        }

        public string FormatUploadCompletedMessage(string operationName, int submittedCellCount)
        {
            var localizedOperationName = LocalizeSyncOperationName(operationName);
            return Locale == "zh"
                ? $"{localizedOperationName}完成。\r\n{SubmittedCellCountLine(submittedCellCount)}"
                : $"{localizedOperationName} completed.\r\n{SubmittedCellCountLine(submittedCellCount)}";
        }

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

        public static bool IsKnownStickyProjectStatus(string text)
        {
            return ForLocale("zh").IsStickyProjectStatus(text) ||
                   ForLocale("en").IsStickyProjectStatus(text);
        }
    }
}
