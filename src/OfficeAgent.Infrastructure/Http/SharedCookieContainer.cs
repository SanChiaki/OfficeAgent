using System.Net;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class SharedCookieContainer
    {
        public CookieContainer Container { get; } = new CookieContainer();
        public string SsoDomain { get; set; } = string.Empty;
    }
}
