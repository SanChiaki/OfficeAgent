using OfficeAgent.Core.Analytics;

namespace OfficeAgent.ExcelAddIn.Analytics
{
    internal sealed class RibbonAnalyticsHelper
    {
        private readonly IAnalyticsService analyticsService;

        public RibbonAnalyticsHelper(IAnalyticsService analyticsService)
        {
            this.analyticsService = analyticsService ?? NoopAnalyticsService.Instance;
        }
    }
}
