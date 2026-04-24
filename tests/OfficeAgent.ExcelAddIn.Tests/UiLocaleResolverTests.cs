using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;
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
            var resolver = new UiLocaleResolver(() => excelUiLocale);

            var resolvedLocale = resolver.Resolve(new AppSettings());

            Assert.Equal(expectedLocale, resolvedLocale);
        }

        [Theory]
        [InlineData("zh", "zh-CN", "zh")]
        [InlineData("en", "zh-CN", "en")]
        [InlineData("ZH", "en-US", "zh")]
        [InlineData("EN", "zh-TW", "en")]
        public void ResolveHonorsExplicitUiLanguageOverrides(string overrideValue, string excelUiLocale, string expectedLocale)
        {
            var resolver = new UiLocaleResolver(() => excelUiLocale);

            var resolvedLocale = resolver.Resolve(new AppSettings
            {
                UiLanguageOverride = overrideValue,
            });

            Assert.Equal(expectedLocale, resolvedLocale);
        }

        [Fact]
        public void ResolveTreatsInvalidOverrideAsSystem()
        {
            var resolver = new UiLocaleResolver(() => "zh-CN");

            var resolvedLocale = resolver.Resolve(new AppSettings
            {
                UiLanguageOverride = "de",
            });

            Assert.Equal("zh", resolvedLocale);
        }
    }
}
