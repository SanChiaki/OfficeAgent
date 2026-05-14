using System;
using System.Threading;
using OfficeAgent.Core.Analytics;

namespace OfficeAgent.ExcelAddIn.Analytics
{
    internal sealed class WorkbookAnalyticsProjectContextProvider : IAnalyticsProjectContextProvider
    {
        private string lastKnownProjectId = string.Empty;

        public WorkbookAnalyticsProjectContextProvider(string initialProjectId = null)
        {
            if (!string.IsNullOrWhiteSpace(initialProjectId))
            {
                lastKnownProjectId = initialProjectId.Trim();
            }
        }

        public string GetCurrentProjectId()
        {
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
    }
}
