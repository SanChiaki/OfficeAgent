using System;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Localization
{
    internal sealed class UiLocaleResolver
    {
        private readonly Func<string> getExcelUiLocale;

        public UiLocaleResolver(Func<string> getExcelUiLocale)
        {
            this.getExcelUiLocale = getExcelUiLocale ?? throw new ArgumentNullException(nameof(getExcelUiLocale));
        }

        public string Resolve(AppSettings settings)
        {
            var normalizedOverride = AppSettings.NormalizeUiLanguageOverride(settings?.UiLanguageOverride);
            if (string.Equals(normalizedOverride, "zh", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalizedOverride, "en", StringComparison.OrdinalIgnoreCase))
            {
                return normalizedOverride;
            }

            var excelUiLocale = getExcelUiLocale() ?? string.Empty;
            return excelUiLocale.StartsWith("zh", StringComparison.OrdinalIgnoreCase) ? "zh" : "en";
        }
    }
}
