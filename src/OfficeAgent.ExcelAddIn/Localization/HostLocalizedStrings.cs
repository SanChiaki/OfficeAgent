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

        public string RibbonAgentButtonLabel => "Open";

        public string RibbonProjectGroupLabel => Locale == "zh" ? "项目" : "Project";

        public string ProjectDropDownPlaceholderText => Locale == "zh" ? "先选择项目" : "Select project";

        public string ProjectDropDownLoginRequiredText => Locale == "zh" ? "请先登录" : "Sign in first";

        public string ProjectDropDownNoAvailableProjectsText => Locale == "zh" ? "无可用项目" : "No projects available";

        public string ProjectDropDownLoadFailedText => Locale == "zh" ? "项目加载失败" : "Failed to load projects";

        public bool IsStickyProjectStatus(string text)
        {
            return string.Equals(text, ProjectDropDownLoginRequiredText, StringComparison.Ordinal) ||
                   string.Equals(text, ProjectDropDownNoAvailableProjectsText, StringComparison.Ordinal) ||
                   string.Equals(text, ProjectDropDownLoadFailedText, StringComparison.Ordinal);
        }

        public string RibbonInitializeSheetButtonLabel => Locale == "zh" ? "初始化当前表" : "Initialize sheet";

        public string RibbonTemplateGroupLabel => Locale == "zh" ? "模板" : "Template";

        public string RibbonApplyTemplateButtonLabel => Locale == "zh" ? "应用模板" : "Apply template";

        public string RibbonSaveTemplateButtonLabel => Locale == "zh" ? "保存模板" : "Save template";

        public string RibbonSaveAsTemplateButtonLabel => Locale == "zh" ? "另存模板" : "Save as template";

        public string RibbonDataSyncGroupLabel => Locale == "zh" ? "数据同步" : "Data sync";

        public string RibbonPartialDownloadButtonLabel => Locale == "zh" ? "部分下载" : "Partial download";

        public string RibbonFullDownloadButtonLabel => Locale == "zh" ? "全量下载" : "Full download";

        public string RibbonPartialUploadButtonLabel => Locale == "zh" ? "部分上传" : "Partial upload";

        public string RibbonFullUploadButtonLabel => Locale == "zh" ? "全量上传" : "Full upload";

        public string RibbonAccountGroupLabel => Locale == "zh" ? "账号" : "Account";

        public string RibbonLoginButtonLabel => Locale == "zh" ? "登录" : "Login";

        public string RibbonLoginInProgressButtonLabel => Locale == "zh" ? "登录中..." : "Signing in...";

        public string RibbonHelpGroupLabel => Locale == "zh" ? "帮助" : "Help";

        public string RibbonDocumentationButtonLabel => Locale == "zh" ? "文档" : "Documentation";

        public string RibbonAboutButtonLabel => Locale == "zh" ? "关于" : "About";

        public string ConfigureSsoUrlFirstMessage => Locale == "zh"
            ? "请先在设置中配置 SSO 地址。"
            : "Configure the SSO URL in Settings first.";

        public string ProjectListEmptyWarningMessage => Locale == "zh"
            ? "项目列表加载完成，但未获取到任何可用项目。\r\n请检查登录状态或项目接口返回。"
            : "The project list loaded, but no projects are available.\r\nCheck your sign-in state or project API response.";

        public string ProjectListLoadFailedMessage(string details)
        {
            return Locale == "zh"
                ? $"项目列表加载失败。\r\n{details}"
                : $"Failed to load projects.\r\n{details}";
        }

        public string DocumentationOpenFailedMessage(string details)
        {
            return Locale == "zh"
                ? $"无法打开文档页面。\r\n{details}"
                : $"Could not open the documentation page.\r\n{details}";
        }

        public string RibbonAboutDialogTitle => Locale == "zh" ? "关于 ISDP" : "About ISDP";

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
