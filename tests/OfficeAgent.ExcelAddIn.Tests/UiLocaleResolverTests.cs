using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UiLocaleResolverTests
    {
        [Theory]
        [InlineData("zh-CN", "zh")]
        [InlineData("zh-TW", "zh")]
        [InlineData("zh-hans", "zh")]
        [InlineData("en-US", "en")]
        [InlineData("ja-JP", "en")]
        [InlineData("", "en")]
        public void ResolveUsesExcelUiLocaleWhenOverrideIsSystem(string excelUiLocale, string expectedLocale)
        {
            var resolver = CreateResolver(excelUiLocale);

            var resolvedLocale = Resolve(resolver, new AppSettings());

            Assert.Equal(expectedLocale, resolvedLocale);
        }

        [Theory]
        [InlineData("zh", "zh-CN", "zh")]
        [InlineData("en", "zh-CN", "en")]
        [InlineData("ZH", "en-US", "zh")]
        [InlineData("EN", "zh-TW", "en")]
        public void ResolveHonorsExplicitUiLanguageOverrides(string overrideValue, string excelUiLocale, string expectedLocale)
        {
            var resolver = CreateResolver(excelUiLocale);

            var resolvedLocale = Resolve(resolver, new AppSettings
            {
                UiLanguageOverride = overrideValue,
            });

            Assert.Equal(expectedLocale, resolvedLocale);
        }

        [Fact]
        public void ResolveTreatsInvalidOverrideAsSystem()
        {
            var resolver = CreateResolver("zh-CN");

            var resolvedLocale = Resolve(resolver, new AppSettings
            {
                UiLanguageOverride = "de",
            });

            Assert.Equal("zh", resolvedLocale);
        }

        private static object CreateResolver(string excelUiLocale)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveRepositoryPath(
                "src",
                "OfficeAgent.ExcelAddIn",
                "bin",
                "Debug",
                "OfficeAgent.ExcelAddIn.dll"));
            var resolverType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Localization.UiLocaleResolver", throwOnError: true);

            return Activator.CreateInstance(resolverType, new object[] { (Func<string>)(() => excelUiLocale) });
        }

        private static string Resolve(object resolver, AppSettings settings)
        {
            var method = resolver.GetType().GetMethod("Resolve", BindingFlags.Instance | BindingFlags.Public);

            Assert.NotNull(method);

            return (string)method.Invoke(resolver, new object[] { settings });
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
