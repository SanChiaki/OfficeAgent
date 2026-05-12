using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Diagnostics;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class AnalyticsServiceTests
    {
        [Fact]
        public void TrackAddsSchemaVersionAndTimestampBeforeWritingToSink()
        {
            var sink = new RecordingAnalyticsSink();
            var service = new AnalyticsService(sink);
            var beforeTrack = DateTime.UtcNow.AddSeconds(-10);

            service.Track(
                "ribbon.download.clicked",
                "ribbon",
                new Dictionary<string, object>
                {
                    { "projectId", "project-123" },
                    { "projectName", "Project Alpha" },
                });

            Assert.True(sink.Written.Wait(TimeSpan.FromSeconds(2)));
            Assert.NotNull(sink.Event);
            Assert.Equal(1, sink.Event.SchemaVersion);
            Assert.Equal("ribbon.download.clicked", sink.Event.EventName);
            Assert.Equal("ribbon", sink.Event.Source);
            Assert.True(sink.Event.OccurredAtUtc >= beforeTrack);
            Assert.True(sink.Event.OccurredAtUtc <= DateTime.UtcNow);
            Assert.Equal("project-123", sink.Event.Properties["projectId"]);
            Assert.Equal("Project Alpha", sink.Event.Properties["projectName"]);
        }

        [Fact]
        public void TrackDoesNotThrowWhenSinkFailsAndWritesDiagnosticLog()
        {
            var entries = new List<OfficeAgentLogEntry>();
            using (var logged = new ManualResetEventSlim(false))
            {
                OfficeAgentLog.Configure(entry =>
                {
                    entries.Add(entry);
                    if (string.Equals(entry.Component, "analytics", StringComparison.Ordinal) &&
                        string.Equals(entry.EventName, "track.failed", StringComparison.Ordinal))
                    {
                        logged.Set();
                    }
                });

                try
                {
                    var service = new AnalyticsService(new FailingAnalyticsSink());

                    var error = Record.Exception(() => service.Track("panel.settings.saved", "panel"));

                    Assert.Null(error);
                    Assert.True(logged.Wait(TimeSpan.FromSeconds(2)));
                    Assert.Contains(
                        entries,
                        entry =>
                            string.Equals(entry.Level, "warn", StringComparison.Ordinal) &&
                            string.Equals(entry.Component, "analytics", StringComparison.Ordinal) &&
                            string.Equals(entry.EventName, "track.failed", StringComparison.Ordinal));
                }
                finally
                {
                    OfficeAgentLog.Reset();
                }
            }
        }

        [Fact]
        public void NoopAnalyticsServiceAcceptsEventsWithoutWriting()
        {
            NoopAnalyticsService.Instance.Track("panel.opened", "panel");
            NoopAnalyticsService.Instance.Track(new AnalyticsEvent
            {
                EventName = "panel.opened",
                Source = "panel",
            });
        }

        private sealed class RecordingAnalyticsSink : IAnalyticsSink
        {
            public ManualResetEventSlim Written { get; } = new ManualResetEventSlim(false);

            public AnalyticsEvent Event { get; private set; }

            public Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
            {
                Event = analyticsEvent;
                Written.Set();
                return Task.CompletedTask;
            }
        }

        private sealed class FailingAnalyticsSink : IAnalyticsSink
        {
            public Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
            {
                throw new InvalidOperationException("Sink failed.");
            }
        }
    }
}
