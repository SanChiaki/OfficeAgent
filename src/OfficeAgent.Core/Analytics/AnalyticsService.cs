using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.Core.Analytics
{
    public sealed class AnalyticsService : IAnalyticsService
    {
        private readonly IAnalyticsSink sink;

        public AnalyticsService(IAnalyticsSink sink)
        {
            this.sink = sink ?? throw new ArgumentNullException(nameof(sink));
        }

        public void Track(AnalyticsEvent analyticsEvent)
        {
            if (analyticsEvent == null || string.IsNullOrWhiteSpace(analyticsEvent.EventName))
            {
                return;
            }

            var normalizedEvent = Normalize(analyticsEvent);

            _ = Task.Run(async () =>
            {
                try
                {
                    await sink.WriteAsync(normalizedEvent, CancellationToken.None).ConfigureAwait(false);
                }
                catch (Exception error)
                {
                    OfficeAgentLog.Warn(
                        "analytics",
                        "track.failed",
                        $"Failed to write analytics event {normalizedEvent.EventName}.",
                        error.ToString());
                }
            });
        }

        public void Track(
            string eventName,
            string source,
            IDictionary<string, object> properties = null,
            IDictionary<string, object> businessContext = null,
            AnalyticsError error = null)
        {
            Track(new AnalyticsEvent
            {
                EventName = eventName,
                Source = source,
                Properties = properties,
                BusinessContext = businessContext,
                Error = error,
            });
        }

        private static AnalyticsEvent Normalize(AnalyticsEvent analyticsEvent)
        {
            analyticsEvent.SchemaVersion = analyticsEvent.SchemaVersion <= 0 ? 1 : analyticsEvent.SchemaVersion;
            analyticsEvent.Source = analyticsEvent.Source ?? string.Empty;
            analyticsEvent.OccurredAtUtc = analyticsEvent.OccurredAtUtc == default(DateTime)
                ? DateTime.UtcNow
                : analyticsEvent.OccurredAtUtc.ToUniversalTime();
            analyticsEvent.Properties = NormalizeDictionary(analyticsEvent.Properties);
            analyticsEvent.BusinessContext = NormalizeDictionary(analyticsEvent.BusinessContext);

            return analyticsEvent;
        }

        private static IDictionary<string, object> NormalizeDictionary(IDictionary<string, object> values)
        {
            var normalized = new Dictionary<string, object>(StringComparer.Ordinal);
            if (values == null)
            {
                return normalized;
            }

            foreach (var value in values)
            {
                normalized[value.Key ?? string.Empty] = value.Value;
            }

            return normalized;
        }
    }
}
