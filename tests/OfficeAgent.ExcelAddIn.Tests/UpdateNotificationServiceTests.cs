using System;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UpdateNotificationServiceTests
    {
        private static readonly DateTime NowUtc = new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc);

        [Fact]
        public async Task CheckForUpdatesAsyncSkipsHttpWhenDisabled()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Disabled(), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.CallCount);
            Assert.False(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public void LoadCachedStateDoesNotShowCachedUpdateWhenDisabled()
        {
            var client = new FakeUpdateManifestClient();
            var store = CreateFreshCachedUpdateStore();
            var service = CreateService(UpdateCheckOptions.Disabled(), client, store);

            service.LoadCachedState();

            Assert.False(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncDoesNotShowCachedUpdateWhenDisabled()
        {
            var client = new FakeUpdateManifestClient();
            var store = CreateFreshCachedUpdateStore();
            var service = CreateService(UpdateCheckOptions.Disabled(), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.CallCount);
            Assert.False(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncDoesNotShowCachedUpdateWhenManifestUrlIsBlank()
        {
            var client = new FakeUpdateManifestClient();
            var store = CreateFreshCachedUpdateStore();
            var service = CreateService(UpdateCheckOptions.Enabled(""), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.CallCount);
            Assert.False(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncUsesCacheWindowWithoutHttp()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    LastCheckedAtUtc = NowUtc.AddHours(-1),
                    LatestVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.CallCount);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.True(service.CurrentState.ShouldShowReminder);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncShowsNewVersionWhenLatestIsHigherAndNotIgnored()
        {
            var publishedAtUtc = new DateTime(2026, 5, 18, 8, 0, 0, DateTimeKind.Utc);
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                    DownloadUrl = "https://updates.example/download.exe",
                    ReleaseNotesUrl = "https://updates.example/notes",
                    PublishedAtUtc = publishedAtUtc,
                    Title = "Release",
                    Summary = "Summary",
                },
            };
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(1, client.CallCount);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.True(service.CurrentState.ShouldShowReminder);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);
            Assert.Equal("https://updates.example/download.exe", service.CurrentState.DownloadUrl);
            Assert.Equal("https://updates.example/notes", service.CurrentState.ReleaseNotesUrl);
            Assert.Equal(publishedAtUtc, service.CurrentState.PublishedAtUtc);
            Assert.Equal("Release", service.CurrentState.Title);
            Assert.Equal("Summary", service.CurrentState.Summary);
            Assert.Equal(NowUtc, store.State.LastCheckedAtUtc);
            Assert.Equal("1.0.176", store.State.LatestVersion);
            Assert.Equal("https://updates.example/download.exe", store.State.DownloadUrl);
            Assert.Equal("https://updates.example/notes", store.State.ReleaseNotesUrl);
            Assert.Equal(publishedAtUtc, store.State.PublishedAtUtc);
            Assert.Equal("Release", store.State.Title);
            Assert.Equal("Summary", store.State.Summary);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncHidesIgnoredReminderButKeepsIgnoredVersionAvailable()
        {
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                },
            };
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    IgnoredVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.True(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);

            client.Manifest = new UpdateManifest
            {
                LatestVersion = "1.0.177",
            };
            store.State.LastCheckedAtUtc = NowUtc.AddDays(-2);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.True(service.CurrentState.HasNewVersion);
            Assert.True(service.CurrentState.ShouldShowReminder);
            Assert.Equal("1.0.177", service.CurrentState.LatestVersion);
            Assert.Equal(2, client.CallCount);
        }

        [Fact]
        public async Task IgnoreCurrentVersionPersistsIgnoredVersionAndRaisesStateChanged()
        {
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                },
            };
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            var stateChangedCount = 0;
            service.StateChanged += (sender, args) => stateChangedCount++;
            await service.CheckForUpdatesAsync(CancellationToken.None);

            service.IgnoreCurrentVersion();

            Assert.Equal("1.0.176", store.State.IgnoredVersion);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
            Assert.Equal("1.0.176", service.CurrentState.LatestVersion);
            Assert.True(stateChangedCount >= 1);
        }

        [Fact]
        public void StartBackgroundCheckReturnsBeforeTheDelayedRequestRuns()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var service = CreateService(
                UpdateCheckOptions.Enabled(
                    "https://updates.example/manifest.json",
                    startupDelay: TimeSpan.FromMilliseconds(200)),
                client,
                store);

            service.StartBackgroundCheck(null);

            Assert.Equal(0, client.CallCount);
        }

        [Fact]
        public void StartBackgroundCheckPostsOneStateChangedToSynchronizationContext()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var context = new RecordingSynchronizationContext();
            var service = CreateService(
                UpdateCheckOptions.Enabled(
                    "https://updates.example/manifest.json",
                    startupDelay: TimeSpan.Zero),
                client,
                store);
            var callerThreadId = Thread.CurrentThread.ManagedThreadId;
            var stateChangedCount = 0;
            var stateChangedThreadId = 0;
            service.StateChanged += (sender, args) =>
            {
                stateChangedCount++;
                stateChangedThreadId = Thread.CurrentThread.ManagedThreadId;
            };

            service.StartBackgroundCheck(context);

            Assert.True(context.WaitForPost(TimeSpan.FromSeconds(5)));
            Assert.Equal(0, stateChangedCount);

            context.RunPostedCallback();

            Assert.Equal(1, stateChangedCount);
            Assert.Equal(1, context.PostCount);
            Assert.Equal(callerThreadId, stateChangedThreadId);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncDoesNotThrowWhenStateChangedSubscriberThrows()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            service.StateChanged += (sender, args) => throw new InvalidOperationException("subscriber failed");

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.True(service.CurrentState.HasNewVersion);
            Assert.True(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncInvokesLaterStateChangedSubscribersWhenEarlierSubscriberThrows()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            var laterSubscriberCallCount = 0;
            service.StateChanged += (sender, args) => throw new InvalidOperationException("subscriber failed");
            service.StateChanged += (sender, args) => laterSubscriberCallCount++;

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(1, laterSubscriberCallCount);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncDoesNotThrowWhenStoreLoadFails()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore
            {
                ThrowOnLoad = true,
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(1, client.CallCount);
            Assert.True(service.CurrentState.HasNewVersion);
        }

        [Fact]
        public async Task IgnoreCurrentVersionDoesNotThrowWhenStoreSaveOrSubscriberFails()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore();
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            await service.CheckForUpdatesAsync(CancellationToken.None);
            store.ThrowOnSave = true;
            service.StateChanged += (sender, args) => throw new InvalidOperationException("subscriber failed");

            service.IgnoreCurrentVersion();

            Assert.True(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncKeepsIgnoredVersionHiddenWhenIgnoreSaveFailsAndCacheIsFresh()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    LastCheckedAtUtc = NowUtc.AddHours(-1),
                    LatestVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            service.LoadCachedState();
            store.ThrowOnSave = true;

            service.IgnoreCurrentVersion();
            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(0, client.CallCount);
            Assert.Equal(string.Empty, store.State.IgnoredVersion);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncPreservesPreviousStateWhenManifestSaveFails()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore
            {
                ThrowOnSave = true,
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(1, client.CallCount);
            Assert.False(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
            Assert.Equal(string.Empty, store.State.LatestVersion);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncDoesNotUseFutureCacheTimestamp()
        {
            var client = new FakeUpdateManifestClient();
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    LastCheckedAtUtc = NowUtc.AddHours(1),
                    LatestVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);

            await service.CheckForUpdatesAsync(CancellationToken.None);

            Assert.Equal(1, client.CallCount);
        }

        [Fact]
        public void UpdateCheckConfigurationCreateDefaultReturnsOptions()
        {
            var options = UpdateCheckConfiguration.CreateDefault();

            Assert.NotNull(options);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncKeepsIgnoredVersionHiddenWhenIgnoredDuringInFlightRequest()
        {
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                },
                WaitForRelease = true,
            };
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    LastCheckedAtUtc = NowUtc.AddDays(-2),
                    LatestVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            service.LoadCachedState();

            var checkTask = service.CheckForUpdatesAsync(CancellationToken.None);
            Assert.True(client.WaitForCall(TimeSpan.FromSeconds(5)));

            service.IgnoreCurrentVersion();
            client.Release();
            await checkTask;

            Assert.Equal("1.0.176", store.State.IgnoredVersion);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        [Fact]
        public async Task CheckForUpdatesAsyncKeepsIgnoredVersionHiddenWhenInFlightManifestSaveFails()
        {
            var client = new FakeUpdateManifestClient
            {
                Manifest = new UpdateManifest
                {
                    LatestVersion = "1.0.176",
                },
                WaitForRelease = true,
            };
            var store = new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    LastCheckedAtUtc = NowUtc.AddDays(-2),
                    LatestVersion = "1.0.176",
                },
            };
            var service = CreateService(UpdateCheckOptions.Enabled("https://updates.example/manifest.json"), client, store);
            service.LoadCachedState();

            var checkTask = service.CheckForUpdatesAsync(CancellationToken.None);
            Assert.True(client.WaitForCall(TimeSpan.FromSeconds(5)));

            service.IgnoreCurrentVersion();
            store.ThrowOnSave = true;
            client.Release();
            await checkTask;

            Assert.Equal("1.0.176", store.State.IgnoredVersion);
            Assert.True(service.CurrentState.HasNewVersion);
            Assert.False(service.CurrentState.ShouldShowReminder);
        }

        private static UpdateNotificationService CreateService(UpdateCheckOptions options, FakeUpdateManifestClient client, IUpdateStateStore store)
        {
            return new UpdateNotificationService(options, client, store, "1.0.175", () => NowUtc);
        }

        private static MemoryUpdateStateStore CreateFreshCachedUpdateStore()
        {
            return new MemoryUpdateStateStore
            {
                State = new UpdateState
                {
                    LastCheckedAtUtc = NowUtc.AddHours(-1),
                    LatestVersion = "1.0.176",
                },
            };
        }

        private sealed class FakeUpdateManifestClient : IUpdateManifestClient
        {
            public UpdateManifest Manifest { get; set; } = new UpdateManifest
            {
                LatestVersion = "1.0.176",
            };

            public int CallCount { get; private set; }
            public bool WaitForRelease { get; set; }
            private readonly ManualResetEventSlim called = new ManualResetEventSlim(false);
            private readonly TaskCompletionSource<bool> released = new TaskCompletionSource<bool>();

            public async Task<UpdateManifest> GetManifestAsync(string manifestUrl, CancellationToken cancellationToken)
            {
                CallCount++;
                called.Set();
                if (WaitForRelease)
                {
                    using (cancellationToken.Register(() => released.TrySetCanceled()))
                    {
                        await released.Task.ConfigureAwait(false);
                    }
                }

                return Manifest;
            }

            public bool WaitForCall(TimeSpan timeout)
            {
                return called.Wait(timeout);
            }

            public void Release()
            {
                released.TrySetResult(true);
            }
        }

        private sealed class MemoryUpdateStateStore : IUpdateStateStore
        {
            public UpdateState State { get; set; } = new UpdateState();
            public bool ThrowOnLoad { get; set; }
            public bool ThrowOnSave { get; set; }

            public UpdateState Load()
            {
                if (ThrowOnLoad)
                {
                    throw new InvalidOperationException("load failed");
                }

                return Clone(State);
            }

            public void Save(UpdateState state)
            {
                if (ThrowOnSave)
                {
                    throw new InvalidOperationException("save failed");
                }

                State = Clone(state);
            }

            private static UpdateState Clone(UpdateState state)
            {
                return new UpdateState
                {
                    LastCheckedAtUtc = state.LastCheckedAtUtc,
                    LatestVersion = state.LatestVersion,
                    DownloadUrl = state.DownloadUrl,
                    ReleaseNotesUrl = state.ReleaseNotesUrl,
                    PublishedAtUtc = state.PublishedAtUtc,
                    Title = state.Title,
                    Summary = state.Summary,
                    IgnoredVersion = state.IgnoredVersion,
                };
            }
        }

        private sealed class RecordingSynchronizationContext : SynchronizationContext
        {
            private readonly ManualResetEventSlim posted = new ManualResetEventSlim(false);
            private SendOrPostCallback callback;
            private object state;

            public int PostCount { get; private set; }

            public override void Post(SendOrPostCallback d, object state)
            {
                callback = d;
                this.state = state;
                PostCount++;
                posted.Set();
            }

            public bool WaitForPost(TimeSpan timeout)
            {
                return posted.Wait(timeout);
            }

            public void RunPostedCallback()
            {
                callback(state);
            }
        }
    }
}
