namespace OfficeAgent.Core.Analytics
{
    public interface IAnalyticsProjectContextProvider
    {
        string GetCurrentProjectId();

        void RememberProjectId(string projectId);
    }
}
