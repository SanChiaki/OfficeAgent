using System;
using System.Threading;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.ExcelAddIn.Analytics
{
    internal sealed class WorkbookAnalyticsProjectContextProvider : IAnalyticsProjectContextProvider
    {
        private readonly Func<string> activeSheetNameProvider;
        private readonly IWorksheetMetadataStore metadataStore;
        private string lastKnownProjectId = string.Empty;

        public WorkbookAnalyticsProjectContextProvider(
            Func<string> activeSheetNameProvider,
            IWorksheetMetadataStore metadataStore)
        {
            this.activeSheetNameProvider = activeSheetNameProvider ?? throw new ArgumentNullException(nameof(activeSheetNameProvider));
            this.metadataStore = metadataStore ?? throw new ArgumentNullException(nameof(metadataStore));
        }

        public string GetCurrentProjectId()
        {
            var currentProjectId = TryGetActiveSheetProjectId();
            if (!string.IsNullOrWhiteSpace(currentProjectId))
            {
                RememberProjectId(currentProjectId);
                return currentProjectId;
            }

            return Volatile.Read(ref lastKnownProjectId) ?? string.Empty;
        }

        public void RememberProjectId(string projectId)
        {
            var normalizedProjectId = (projectId ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedProjectId))
            {
                return;
            }

            Volatile.Write(ref lastKnownProjectId, normalizedProjectId);
        }

        private string TryGetActiveSheetProjectId()
        {
            try
            {
                var sheetName = activeSheetNameProvider.Invoke() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    return string.Empty;
                }

                var binding = metadataStore.LoadBinding(sheetName);
                return binding?.ProjectId ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
