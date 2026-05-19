using System;
using Microsoft.Win32;
using OfficeAgent.Core.Diagnostics;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal static class UpdateCheckConfiguration
    {
        private const string RegistryPath = @"Software\OfficeAgent";
        private const string ManifestUrlValueName = "UpdateManifestUrl";
        private const string DefaultManifestUrl = "";

        public static UpdateCheckOptions CreateDefault()
        {
#if DEBUG
            OfficeAgentLog.Info("updates", "configuration.disabled_debug", "Update checks are disabled in Debug builds.");
            return UpdateCheckOptions.Disabled();
#else
            var manifestUrl = ReadManifestUrl(Registry.CurrentUser);
            if (string.IsNullOrWhiteSpace(manifestUrl))
            {
                manifestUrl = ReadManifestUrl(Registry.LocalMachine);
            }

            if (string.IsNullOrWhiteSpace(manifestUrl))
            {
                manifestUrl = DefaultManifestUrl;
            }

            if (string.IsNullOrWhiteSpace(manifestUrl))
            {
                OfficeAgentLog.Info("updates", "configuration.disabled_missing_url", "Update checks are disabled because no manifest URL is configured.");
                return UpdateCheckOptions.Disabled();
            }

            return UpdateCheckOptions.Enabled(manifestUrl.Trim());
#endif
        }

        public static UpdateCheckOptions Load()
        {
            return CreateDefault();
        }

        private static string ReadManifestUrl(RegistryKey root)
        {
            try
            {
                using (var key = root.OpenSubKey(RegistryPath))
                {
                    return key?.GetValue(ManifestUrlValueName) as string ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                OfficeAgentLog.Warn("updates", "configuration.registry_read_failed", "Failed to read update manifest URL from the registry.", ex.Message);
                return string.Empty;
            }
        }
    }
}
