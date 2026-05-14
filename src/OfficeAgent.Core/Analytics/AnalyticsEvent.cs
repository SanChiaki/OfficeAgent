using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace OfficeAgent.Core.Analytics
{
    public sealed class AnalyticsEvent
    {
        public int SchemaVersion { get; set; } = 1;

        public string Version { get; set; } = string.Empty;

        public string EventName { get; set; } = string.Empty;

        public string Source { get; set; } = string.Empty;

        public DateTime OccurredAtUtc { get; set; }

        public IDictionary<string, object> Properties { get; set; } = new Dictionary<string, object>(StringComparer.Ordinal);

        public IDictionary<string, object> BusinessContext { get; set; } = new Dictionary<string, object>(StringComparer.Ordinal);

        public AnalyticsError Error { get; set; }

        [JsonIgnore]
        public string EnvelopeProjectId { get; set; } = string.Empty;
    }
}
