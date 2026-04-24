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
            var type = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings",
                throwOnError: true);
            var method = type.GetMethod("ForLocale", BindingFlags.Public | BindingFlags.Static);

            Assert.NotNull(method);

            return method.Invoke(null, new object[] { locale });
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
