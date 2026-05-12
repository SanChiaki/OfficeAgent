namespace OfficeAgent.Core.Analytics
{
    public sealed class AnalyticsError
    {
        public string Code { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public string ExceptionType { get; set; } = string.Empty;
    }
}
