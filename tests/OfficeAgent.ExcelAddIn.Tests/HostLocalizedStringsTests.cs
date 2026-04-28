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
        [InlineData("zh", "项目", "先选择项目", "初始化当前表", "数据同步", "文档", "关于")]
        [InlineData("en", "Project", "Select project", "Initialize sheet", "Data sync", "Documentation", "About")]
        public void ForLocaleReturnsExpectedRibbonLabels(
            string locale,
            string expectedProjectGroup,
            string expectedProjectPlaceholder,
            string expectedInitializeSheet,
            string expectedDataSyncGroup,
            string expectedDocumentation,
            string expectedAbout)
        {
            var strings = CreateStrings(locale);

            Assert.Equal(expectedProjectGroup, GetString(strings, "RibbonProjectGroupLabel"));
            Assert.Equal(expectedProjectPlaceholder, GetString(strings, "ProjectDropDownPlaceholderText"));
            Assert.Equal(expectedInitializeSheet, GetString(strings, "RibbonInitializeSheetButtonLabel"));
            Assert.Equal(expectedDataSyncGroup, GetString(strings, "RibbonDataSyncGroupLabel"));
            Assert.Equal(expectedDocumentation, GetString(strings, "RibbonDocumentationButtonLabel"));
            Assert.Equal(expectedAbout, GetString(strings, "RibbonAboutButtonLabel"));
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
        [InlineData("", "en")]
        [InlineData("de", "en")]
        [InlineData("zh-CN", "en")]
        [InlineData("ZH", "zh")]
        public void ForLocaleNormalizesUnsupportedLocalesToSupportedSet(string requestedLocale, string expectedLocale)
        {
            var strings = CreateStrings(requestedLocale);

            Assert.Equal(expectedLocale, GetString(strings, "Locale"));
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
