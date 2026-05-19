using System;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateNotificationService
    {
        private readonly object syncRoot = new object();
        private readonly UpdateCheckOptions options;
        private readonly IUpdateManifestClient manifestClient;
        private readonly IUpdateStateStore stateStore;
        private readonly string currentVersion;
        private readonly Func<DateTime> getUtcNow;
        private UpdateState cachedState = new UpdateState();

        public UpdateNotificationService(
            UpdateCheckOptions options,
            IUpdateManifestClient manifestClient,
            IUpdateStateStore stateStore,
            string currentVersion,
            Func<DateTime> getUtcNow = null)
        {
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.manifestClient = manifestClient ?? throw new ArgumentNullException(nameof(manifestClient));
            this.stateStore = stateStore ?? throw new ArgumentNullException(nameof(stateStore));
            this.currentVersion = currentVersion ?? string.Empty;
            this.getUtcNow = getUtcNow ?? (() => DateTime.UtcNow);
            CurrentState = UpdateNotificationState.Empty;
        }

        public event EventHandler StateChanged;

        public UpdateNotificationState CurrentState { get; private set; }

        public void LoadCachedState()
        {
            ApplyState(stateStore.Load(), raiseStateChanged: true);
        }

        public void StartBackgroundCheck(SynchronizationContext uiContext)
        {
            Task.Run(async () =>
            {
                try
                {
                    if (options.StartupDelay > TimeSpan.Zero)
                    {
                        await Task.Delay(options.StartupDelay).ConfigureAwait(false);
                    }

                    await CheckForUpdatesAsync(CancellationToken.None).ConfigureAwait(false);
                    RaiseStateChanged(uiContext);
                }
                catch (Exception ex)
                {
                    OfficeAgentLog.Error("updates", "background_check.failed", "Background update check failed.", ex);
                }
            });
        }

        public async Task CheckForUpdatesAsync(CancellationToken cancellationToken)
        {
            var previousState = stateStore.Load();

            try
            {
                if (!options.IsEnabled || string.IsNullOrWhiteSpace(options.ManifestUrl))
                {
                    OfficeAgentLog.Info("updates", "check.skipped_disabled", "Update check skipped because update checks are disabled.");
                    ApplyState(previousState, raiseStateChanged: true);
                    return;
                }

                var nowUtc = getUtcNow();
                if (IsCacheFresh(previousState, nowUtc))
                {
                    ApplyState(previousState, raiseStateChanged: true);
                    return;
                }

                using (var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken))
                {
                    timeoutCts.CancelAfter(options.RequestTimeout);
                    var manifest = await manifestClient.GetManifestAsync(options.ManifestUrl, timeoutCts.Token).ConfigureAwait(false);
                    if (!UpdateVersionComparer.TryParse(manifest?.LatestVersion, out var latestVersion))
                    {
                        throw new InvalidOperationException("Update manifest latestVersion is invalid.");
                    }

                    var nextState = CloneState(previousState);
                    nextState.ApplyManifest(manifest, nowUtc);
                    stateStore.Save(nextState);
                    ApplyState(nextState, raiseStateChanged: true);
                }
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "check.failed", "Update check failed.", ex.Message);
                ApplyState(previousState, raiseStateChanged: true);
            }
        }

        public void IgnoreCurrentVersion()
        {
            UpdateState nextState;
            lock (syncRoot)
            {
                if (string.IsNullOrWhiteSpace(cachedState.LatestVersion))
                {
                    return;
                }

                nextState = CloneState(cachedState);
                nextState.IgnoredVersion = nextState.LatestVersion;
                cachedState = nextState;
                CurrentState = BuildNotificationState(nextState);
            }

            stateStore.Save(nextState);
            OfficeAgentLog.Info("updates", "version.ignored", "User ignored the current update version.", nextState.IgnoredVersion);
            RaiseStateChanged(null);
        }

        private bool IsCacheFresh(UpdateState state, DateTime nowUtc)
        {
            return state.LastCheckedAtUtc.HasValue &&
                   nowUtc - state.LastCheckedAtUtc.Value < options.CacheDuration;
        }

        private void ApplyState(UpdateState state, bool raiseStateChanged)
        {
            lock (syncRoot)
            {
                cachedState = CloneState(state);
                CurrentState = BuildNotificationState(cachedState);
            }

            if (raiseStateChanged)
            {
                RaiseStateChanged(null);
            }
        }

        private UpdateNotificationState BuildNotificationState(UpdateState state)
        {
            if (state == null ||
                !UpdateVersionComparer.IsNewerThanCurrent(state.LatestVersion, currentVersion) ||
                string.Equals(state.LatestVersion, state.IgnoredVersion, StringComparison.OrdinalIgnoreCase))
            {
                return UpdateNotificationState.Empty;
            }

            return new UpdateNotificationState(
                true,
                state.LatestVersion,
                state.DownloadUrl,
                state.ReleaseNotesUrl,
                state.PublishedAtUtc,
                state.Title,
                state.Summary);
        }

        private void RaiseStateChanged(SynchronizationContext uiContext)
        {
            var handler = StateChanged;
            if (handler == null)
            {
                return;
            }

            if (uiContext == null)
            {
                handler(this, EventArgs.Empty);
                return;
            }

            uiContext.Post(_ => handler(this, EventArgs.Empty), null);
        }

        private static UpdateState CloneState(UpdateState state)
        {
            state = state ?? new UpdateState();
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
}
