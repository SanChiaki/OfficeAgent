using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class HostLocalizedStringsTests
    {
        [Theory]
        [InlineData("zh", "先选择项目", "请先登录", "配置当前表布局", "点我登录")]
        [InlineData("en", "Select project", "Sign in first", "Configure sheet layout", "Sign in")]
        public void ForLocaleReturnsExpectedLocalizedChrome(
            string locale,
            string expectedPlaceholder,
            string expectedLoginRequired,
            string expectedLayoutTitle,
            string expectedSignInButton)
        {
            var strings = CreateStrings(locale);

            Assert.Equal(expectedPlaceholder, GetString(strings, "ProjectDropDownPlaceholderText"));
            Assert.Equal(expectedLoginRequired, GetString(strings, "ProjectDropDownLoginRequiredText"));
            Assert.Equal(expectedLayoutTitle, GetString(strings, "ProjectLayoutDialogTitle"));
            Assert.Equal(expectedSignInButton, GetString(strings, "AuthenticationRequiredLoginButtonText"));
        }

        [Theory]
        [InlineData("zh", "下面三个值会写入当前工作表与ISDP实施计划的映射配置表xISDP_Setting中，请确认后保存。")]
        [InlineData("en", "The three values below will be written to xISDP_Setting, the mapping configuration table for the current worksheet and the ISDP implementation plan. Confirm them before saving.")]
        public void ForLocaleReturnsExpectedProjectLayoutInstruction(
            string locale,
            string expectedInstruction)
        {
            var strings = CreateStrings(locale);

            Assert.Equal(expectedInstruction, GetString(strings, "ProjectLayoutInstructionText"));
        }

        [Theory]
        [InlineData("zh")]
        [InlineData("en")]
        public void ForLocaleUsesXisdpAsHostAndRibbonName(string locale)
        {
            var strings = CreateStrings(locale);

            Assert.Equal("xISDP", GetString(strings, "HostWindowTitle"));
            Assert.Equal("xISDP", GetString(strings, "RibbonTabLabel"));
            Assert.Equal("xISDP AI", GetString(strings, "RibbonAgentGroupLabel"));
            Assert.Equal(string.Empty, GetString(strings, "RibbonAgentButtonLabel"));
        }

        [Theory]
        [InlineData("zh", "下载", "上传")]
        [InlineData("en", "Download", "Upload")]
        public void ForLocaleReturnsExpectedRibbonSyncButtonLabels(
            string locale,
            string expectedDownloadLabel,
            string expectedUploadLabel)
        {
            var strings = CreateStrings(locale);

            Assert.Equal(expectedDownloadLabel, GetString(strings, "RibbonPartialDownloadButtonLabel"));
            Assert.Equal(expectedUploadLabel, GetString(strings, "RibbonPartialUploadButtonLabel"));
        }

        [Theory]
        [InlineData("zh", "AI映射列")]
        [InlineData("en", "AI map columns")]
        public void ForLocaleReturnsAiColumnMappingRibbonLabel(string locale, string expectedLabel)
        {
            var strings = CreateStrings(locale);

            Assert.Equal(expectedLabel, GetString(strings, "RibbonAiMapColumnsButtonLabel"));
        }

        [Theory]
        [InlineData("zh", "配置", "应用配置", "保存配置", "另存配置")]
        [InlineData("en", "Setting", "Apply Setting", "Save Setting", "Save as Setting")]
        public void ForLocaleReturnsExpectedConfigurationRibbonLabels(
            string locale,
            string expectedGroupLabel,
            string expectedApplyLabel,
            string expectedSaveLabel,
            string expectedSaveAsLabel)
        {
            var strings = CreateStrings(locale);

            Assert.Equal(expectedGroupLabel, GetString(strings, "RibbonTemplateGroupLabel"));
            Assert.Equal(expectedApplyLabel, GetString(strings, "RibbonApplyTemplateButtonLabel"));
            Assert.Equal(expectedSaveLabel, GetString(strings, "RibbonSaveTemplateButtonLabel"));
            Assert.Equal(expectedSaveAsLabel, GetString(strings, "RibbonSaveAsTemplateButtonLabel"));
        }

        [Theory]
        [InlineData("", "en")]
        [InlineData("de", "en")]
        [InlineData("zh-CN", "en")]
        [InlineData("ZH", "zh")]
        public void ForLocaleNormalizesUnsupportedLocalesToSupportedSet(string requestedLocale, string expectedLocale)
        {
            var strings = CreateStrings(requestedLocale);

            Assert.Equal(expectedLocale, GetString(strings, "Locale"));
        }

        [Theory]
        [InlineData("zh", "请先登录", true)]
        [InlineData("zh", "无可用项目", true)]
        [InlineData("zh", "项目加载失败", true)]
        [InlineData("zh", "先选择项目", false)]
        [InlineData("en", "Sign in first", true)]
        [InlineData("en", "No projects available", true)]
        [InlineData("en", "Failed to load projects", true)]
        [InlineData("en", "Select project", false)]
        public void IsStickyProjectStatusMatchesLocalizedStatusPolicy(string locale, string text, bool expected)
        {
            var strings = CreateStrings(locale);
            var method = strings.GetType().GetMethod("IsStickyProjectStatus", BindingFlags.Instance | BindingFlags.Public);

            Assert.NotNull(method);
            Assert.Equal(expected, (bool)method.Invoke(strings, new object[] { text }));
        }

        [Theory]
        [InlineData("请先登录", true)]
        [InlineData("No projects available", true)]
        [InlineData("先选择项目", false)]
        [InlineData("Select project", false)]
        [InlineData("random", false)]
        public void IsKnownStickyProjectStatusRecognizesCanonicalStatusesAcrossLocales(string text, bool expected)
        {
            var type = GetStringsType();
            var method = type.GetMethod("IsKnownStickyProjectStatus", BindingFlags.Public | BindingFlags.Static);

            Assert.NotNull(method);
            Assert.Equal(expected, (bool)method.Invoke(null, new object[] { text }));
        }

        [Theory]
        [InlineData("zh", "全量下载", "全量下载", "全量下载完成。\r\n记录数：3\r\n字段数：4", "查询结果为空，请确认列名是否正确匹配。", "全量上传没有可提交的单元格。", "全量上传完成。\r\n提交单元格数：2")]
        [InlineData("en", "全量下载", "Full download", "Full download completed.\r\nRows: 3\r\nFields: 4", "The query result is empty. Check whether the column names are mapped correctly.", "Full upload has no cells to submit.", "Full upload completed.\r\nSubmitted cells: 2")]
        public void ForLocaleFormatsSyncOperationMessages(
            string locale,
            string operationName,
            string expectedLocalizedOperationName,
            string expectedDownloadCompletedMessage,
            string expectedDownloadNoMatchingRowsMessage,
            string expectedUploadNoChangesMessage,
            string expectedUploadCompletedMessage)
        {
            var strings = CreateStrings(locale);
            var localizeMethod = strings.GetType().GetMethod("LocalizeSyncOperationName", BindingFlags.Instance | BindingFlags.Public);
            var downloadCompletedMethod = strings.GetType().GetMethod("FormatDownloadCompletedMessage", BindingFlags.Instance | BindingFlags.Public);
            var downloadNoMatchingRowsMethod = strings.GetType().GetMethod("FormatDownloadNoMatchingRowsMessage", BindingFlags.Instance | BindingFlags.Public);
            var uploadNoChangesMethod = strings.GetType().GetMethod("FormatUploadNoChangesMessage", BindingFlags.Instance | BindingFlags.Public);
            var uploadCompletedMethod = strings.GetType().GetMethod("FormatUploadCompletedMessage", BindingFlags.Instance | BindingFlags.Public);

            Assert.NotNull(localizeMethod);
            Assert.NotNull(downloadCompletedMethod);
            Assert.NotNull(downloadNoMatchingRowsMethod);
            Assert.NotNull(uploadNoChangesMethod);
            Assert.NotNull(uploadCompletedMethod);

            Assert.Equal(expectedLocalizedOperationName, (string)localizeMethod.Invoke(strings, new object[] { operationName }));
            Assert.Equal(expectedDownloadCompletedMessage, (string)downloadCompletedMethod.Invoke(strings, new object[] { operationName, 3, 4 }));
            Assert.Equal(expectedDownloadNoMatchingRowsMessage, (string)downloadNoMatchingRowsMethod.Invoke(strings, new object[] { operationName }));
            Assert.Equal(expectedUploadNoChangesMessage, (string)uploadNoChangesMethod.Invoke(strings, new object[] { "全量上传" }));
            Assert.Equal(expectedUploadCompletedMessage, (string)uploadCompletedMethod.Invoke(strings, new object[] { "全量上传", 2 }));
        }

        [Theory]
        [InlineData("zh", "部分下载", "下载")]
        [InlineData("zh", "部分上传", "上传")]
        [InlineData("en", "部分下载", "Download")]
        [InlineData("en", "部分上传", "Upload")]
        public void ForLocaleLocalizesPartialSyncOperationsToVisibleButtonLabels(
            string locale,
            string operationName,
            string expectedLocalizedOperationName)
        {
            var strings = CreateStrings(locale);
            var localizeMethod = strings.GetType().GetMethod("LocalizeSyncOperationName", BindingFlags.Instance | BindingFlags.Public);

            Assert.NotNull(localizeMethod);
            Assert.Equal(expectedLocalizedOperationName, (string)localizeMethod.Invoke(strings, new object[] { operationName }));
        }

        private static object CreateStrings(string locale)
        {
            var type = GetStringsType();
            var method = type.GetMethod("ForLocale", BindingFlags.Public | BindingFlags.Static);

            Assert.NotNull(method);

            return method.Invoke(null, new object[] { locale });
        }

        private static Type GetStringsType()
        {
            return LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings",
                throwOnError: true);
        }

        private static string GetString(object instance, string propertyName)
        {
            var property = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);

            Assert.NotNull(property);

            return (string)property.GetValue(instance);
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
