using System;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal sealed class UpdateCheckOptions
    {
        private static readonly TimeSpan DefaultCacheDuration = TimeSpan.FromHours(24);
        private static readonly TimeSpan DefaultRequestTimeout = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan DefaultStartupDelay = TimeSpan.FromSeconds(5);

        private UpdateCheckOptions(bool isEnabled, string manifestUrl, TimeSpan cacheDuration, TimeSpan requestTimeout, TimeSpan startupDelay)
        {
            IsEnabled = isEnabled;
            ManifestUrl = manifestUrl ?? string.Empty;
            CacheDuration = cacheDuration;
            RequestTimeout = requestTimeout;
            StartupDelay = startupDelay;
        }

        public bool IsEnabled { get; }
        public string ManifestUrl { get; }
        public TimeSpan CacheDuration { get; }
        public TimeSpan RequestTimeout { get; }
        public TimeSpan StartupDelay { get; }

        public static UpdateCheckOptions Enabled(string manifestUrl, TimeSpan? cacheDuration = null, TimeSpan? requestTimeout = null, TimeSpan? startupDelay = null)
        {
            return new UpdateCheckOptions(
                true,
                manifestUrl,
                cacheDuration ?? DefaultCacheDuration,
                requestTimeout ?? DefaultRequestTimeout,
                startupDelay ?? DefaultStartupDelay);
        }

        public static UpdateCheckOptions Disabled()
        {
            return new UpdateCheckOptions(false, string.Empty, DefaultCacheDuration, DefaultRequestTimeout, DefaultStartupDelay);
        }
    }
}
