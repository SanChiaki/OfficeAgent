using System.Collections.Generic;

namespace OfficeAgent.Core.Analytics
{
    public interface IAnalyticsService
    {
        void Track(AnalyticsEvent analyticsEvent);

        void Track(
            string eventName,
            string source,
            IDictionary<string, object> properties = null,
            IDictionary<string, object> businessContext = null,
            AnalyticsError error = null);
    }
}
