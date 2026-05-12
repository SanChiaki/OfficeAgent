using System;
using System.Collections.Generic;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;
using OfficeAgent.ExcelAddIn.Localization;

namespace OfficeAgent.ExcelAddIn.Analytics
{
    internal sealed class RibbonAnalyticsHelper
    {
        private readonly IAnalyticsService analyticsService;
        private readonly Func<SheetBinding> getActiveBinding;
        private readonly Func<string> getActiveSheetName;
        private readonly Func<string> getActiveWorkbookName;
        private readonly Func<HostLocalizedStrings> getStrings;

        public RibbonAnalyticsHelper(
            IAnalyticsService analyticsService,
            Func<SheetBinding> getActiveBinding,
            Func<string> getActiveSheetName,
            Func<string> getActiveWorkbookName,
            Func<HostLocalizedStrings> getStrings)
        {
            this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;
            this.getActiveBinding = getActiveBinding ?? (() => null);
            this.getActiveSheetName = getActiveSheetName ?? (() => string.Empty);
            this.getActiveWorkbookName = getActiveWorkbookName ?? (() => string.Empty);
            this.getStrings = getStrings ?? (() => HostLocalizedStrings.ForLocale("en"));
        }

        public void Track(
            string eventName,
            IDictionary<string, object> properties = null,
            AnalyticsError error = null)
        {
            if (string.IsNullOrWhiteSpace(eventName))
            {
                return;
            }

            var binding = SafeInvoke(getActiveBinding);
            var strings = SafeInvoke(getStrings) ?? HostLocalizedStrings.ForLocale("en");
            var merged = new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["systemKey"] = binding?.SystemKey ?? string.Empty,
                ["projectId"] = binding?.ProjectId ?? string.Empty,
                ["projectName"] = binding?.ProjectName ?? string.Empty,
                ["sheetName"] = SafeInvoke(getActiveSheetName) ?? string.Empty,
                ["workbookName"] = SafeInvoke(getActiveWorkbookName) ?? string.Empty,
                ["uiLocale"] = strings.Locale ?? string.Empty,
            };

            if (properties != null)
            {
                foreach (var property in properties)
                {
                    merged[property.Key ?? string.Empty] = property.Value;
                }
            }

            analyticsService.Track(eventName, "ribbon", merged, error: error);
        }

        private static T SafeInvoke<T>(Func<T> valueProvider)
        {
            try
            {
                return valueProvider == null ? default(T) : valueProvider();
            }
            catch
            {
                return default(T);
            }
        }
    }
}
