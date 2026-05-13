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
        private readonly string defaultVersion;

        public AnalyticsService(IAnalyticsSink sink, string defaultVersion = null)
        {
            this.sink = sink ?? throw new ArgumentNullException(nameof(sink));
            this.defaultVersion = defaultVersion ?? string.Empty;
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
            return new AnalyticsEvent
            {
                SchemaVersion = analyticsEvent.SchemaVersion <= 0 ? 1 : analyticsEvent.SchemaVersion,
                Version = NormalizeVersion(analyticsEvent.Version, defaultVersion),
                EventName = analyticsEvent.EventName,
                Source = analyticsEvent.Source ?? string.Empty,
                OccurredAtUtc = NormalizeTimestamp(analyticsEvent.OccurredAtUtc),
                Properties = CopyDictionary(analyticsEvent.Properties),
                BusinessContext = CopyDictionary(analyticsEvent.BusinessContext),
                Error = CopyError(analyticsEvent.Error),
            };
        }

        private static DateTime NormalizeTimestamp(DateTime occurredAtUtc)
        {
            if (occurredAtUtc == default(DateTime))
            {
                return DateTime.UtcNow;
            }

            return occurredAtUtc.Kind == DateTimeKind.Local
                ? occurredAtUtc.ToUniversalTime()
                : occurredAtUtc;
        }

        private static string NormalizeVersion(string eventVersion, string defaultVersion)
        {
            if (!string.IsNullOrWhiteSpace(eventVersion))
            {
                return eventVersion.Trim();
            }

            return string.IsNullOrWhiteSpace(defaultVersion) ? string.Empty : defaultVersion.Trim();
        }

        private static IDictionary<string, object> CopyDictionary(IDictionary<string, object> values)
        {
            var copy = new Dictionary<string, object>(StringComparer.Ordinal);
            if (values == null)
            {
                return copy;
            }

            foreach (var value in values)
            {
                copy[value.Key ?? string.Empty] = value.Value;
            }

            return copy;
        }

        private static AnalyticsError CopyError(AnalyticsError error)
        {
            if (error == null)
            {
                return null;
            }

            return new AnalyticsError
            {
                Code = error.Code ?? string.Empty,
                Message = error.Message ?? string.Empty,
                ExceptionType = error.ExceptionType ?? string.Empty,
            };
        }
    }
}
