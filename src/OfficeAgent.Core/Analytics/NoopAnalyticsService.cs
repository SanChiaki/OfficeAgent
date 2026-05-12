using System.Collections.Generic;

namespace OfficeAgent.Core.Analytics
{
    public sealed class NoopAnalyticsService : IAnalyticsService
    {
        public static readonly NoopAnalyticsService Instance = new NoopAnalyticsService();

        private NoopAnalyticsService()
        {
        }

        public void Track(AnalyticsEvent analyticsEvent)
        {
        }

        public void Track(
            string eventName,
            string source,
            IDictionary<string, object> properties = null,
            IDictionary<string, object> businessContext = null,
            AnalyticsError error = null)
        {
        }
    }
}
