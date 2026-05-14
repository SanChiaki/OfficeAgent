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
            var service = new AnalyticsService(sink, "1.0.175");
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
            Assert.Equal("1.0.175", sink.Event.Version);
            Assert.Equal("ribbon.download.clicked", sink.Event.EventName);
            Assert.Equal("ribbon", sink.Event.Source);
            Assert.True(sink.Event.OccurredAtUtc >= beforeTrack);
            Assert.True(sink.Event.OccurredAtUtc <= DateTime.UtcNow);
            Assert.Equal("project-123", sink.Event.Properties["projectId"]);
            Assert.Equal("Project Alpha", sink.Event.Properties["projectName"]);
            Assert.Equal("project-123", sink.Event.EnvelopeProjectId);
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
        public void TrackSnapshotsCallerOwnedAnalyticsEventBeforeAsynchronousWrite()
        {
            var sink = new BlockingAnalyticsSink();
            var service = new AnalyticsService(sink, "1.0.175");
            var occurredAtUtc = new DateTime(2026, 5, 12, 9, 30, 0, DateTimeKind.Unspecified);
            var analyticsEvent = new AnalyticsEvent
            {
                Version = "1.0.174",
                EventName = "panel.opened",
                Source = "panel",
                OccurredAtUtc = occurredAtUtc,
                Properties = new Dictionary<string, object>
                {
                    { "projectId", "project-123" },
                },
                BusinessContext = new Dictionary<string, object>
                {
                    { "tenantId", "tenant-123" },
                },
                Error = new AnalyticsError
                {
                    Code = "original-code",
                    Message = "Original message.",
                    ExceptionType = "OriginalException",
                },
            };

            service.Track(analyticsEvent);

            Assert.True(sink.Entered.Wait(TimeSpan.FromSeconds(2)));
            analyticsEvent.EventName = "panel.mutated";
            analyticsEvent.Source = "mutated";
            analyticsEvent.OccurredAtUtc = occurredAtUtc.AddHours(1);
            analyticsEvent.Properties["projectId"] = "mutated-project";
            analyticsEvent.Properties["newProperty"] = "mutated";
            analyticsEvent.BusinessContext["tenantId"] = "mutated-tenant";
            analyticsEvent.BusinessContext["newContext"] = "mutated";
            analyticsEvent.Error.Code = "mutated-code";
            analyticsEvent.Error.Message = "Mutated message.";
            analyticsEvent.Error.ExceptionType = "MutatedException";

            sink.AllowWrite.Set();

            Assert.True(sink.Written.Wait(TimeSpan.FromSeconds(2)));
            Assert.NotNull(sink.Event);
            Assert.NotSame(analyticsEvent, sink.Event);
            Assert.Equal("1.0.174", sink.Event.Version);
            Assert.Equal("panel.opened", sink.Event.EventName);
            Assert.Equal("panel", sink.Event.Source);
            Assert.Equal(occurredAtUtc, sink.Event.OccurredAtUtc);
            Assert.Equal("project-123", sink.Event.Properties["projectId"]);
            Assert.False(sink.Event.Properties.ContainsKey("newProperty"));
            Assert.Equal("tenant-123", sink.Event.BusinessContext["tenantId"]);
            Assert.False(sink.Event.BusinessContext.ContainsKey("newContext"));
            Assert.NotSame(analyticsEvent.Error, sink.Event.Error);
            Assert.Equal("original-code", sink.Event.Error.Code);
            Assert.Equal("Original message.", sink.Event.Error.Message);
            Assert.Equal("OriginalException", sink.Event.Error.ExceptionType);
        }

        [Fact]
        public void TrackUsesDefaultVersionWhenEventDoesNotProvideOne()
        {
            var sink = new RecordingAnalyticsSink();
            var service = new AnalyticsService(sink, "1.0.175");

            service.Track(new AnalyticsEvent
            {
                EventName = "panel.settings.saved",
                Source = "panel",
            });

            Assert.True(sink.Written.Wait(TimeSpan.FromSeconds(2)));
            Assert.NotNull(sink.Event);
            Assert.Equal("1.0.175", sink.Event.Version);
        }

        [Fact]
        public void TrackSnapshotsFallbackEnvelopeProjectIdFromProvider()
        {
            var sink = new RecordingAnalyticsSink();
            var provider = new RecordingProjectContextProvider("fallback-project-001");
            var service = new AnalyticsService(sink, "1.0.175", provider);

            service.Track(new AnalyticsEvent
            {
                EventName = "panel.settings.saved",
                Source = "panel",
                Properties = new Dictionary<string, object>
                {
                    { "sessionId", "session-123" },
                },
            });

            Assert.True(sink.Written.Wait(TimeSpan.FromSeconds(2)));
            Assert.NotNull(sink.Event);
            Assert.Equal("fallback-project-001", sink.Event.EnvelopeProjectId);
            Assert.Equal(1, provider.GetCurrentProjectIdCallCount);
            Assert.Equal(0, provider.RememberProjectIdCallCount);
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

        private sealed class BlockingAnalyticsSink : IAnalyticsSink
        {
            public ManualResetEventSlim Entered { get; } = new ManualResetEventSlim(false);

            public ManualResetEventSlim AllowWrite { get; } = new ManualResetEventSlim(false);

            public ManualResetEventSlim Written { get; } = new ManualResetEventSlim(false);

            public AnalyticsEvent Event { get; private set; }

            public Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
            {
                Entered.Set();
                AllowWrite.Wait(cancellationToken);
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

        private sealed class RecordingProjectContextProvider : IAnalyticsProjectContextProvider
        {
            private readonly string currentProjectId;

            public RecordingProjectContextProvider(string currentProjectId)
            {
                this.currentProjectId = currentProjectId;
            }

            public int GetCurrentProjectIdCallCount { get; private set; }

            public int RememberProjectIdCallCount { get; private set; }

            public string GetCurrentProjectId()
            {
                GetCurrentProjectIdCallCount++;
                return currentProjectId;
            }

            public void RememberProjectId(string projectId)
            {
                RememberProjectIdCallCount++;
            }
        }
    }
}
