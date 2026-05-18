using System;
using System.Net;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class AccountSessionService
    {
        private readonly SharedCookieContainer sharedCookies;
        private readonly FileCookieStore cookieStore;

        public AccountSessionService(SharedCookieContainer sharedCookies, FileCookieStore cookieStore)
        {
            this.sharedCookies = sharedCookies ?? throw new ArgumentNullException(nameof(sharedCookies));
            this.cookieStore = cookieStore ?? throw new ArgumentNullException(nameof(cookieStore));
        }

        public void ConfigureSsoDomain(string ssoUrl)
        {
            if (string.IsNullOrWhiteSpace(ssoUrl))
            {
                sharedCookies.SsoDomain = string.Empty;
                return;
            }

            try
            {
                sharedCookies.SsoDomain = new Uri(ssoUrl).Host;
            }
            catch (UriFormatException)
            {
                sharedCookies.SsoDomain = string.Empty;
            }
        }

        public bool IsLoggedIn()
        {
            var ssoDomain = sharedCookies.SsoDomain;
            if (string.IsNullOrWhiteSpace(ssoDomain))
            {
                return false;
            }

            try
            {
                return sharedCookies.Container.GetCookies(CreateSsoCookieUri(ssoDomain)).Count > 0;
            }
            catch (UriFormatException)
            {
                return false;
            }
        }

        public void Logout()
        {
            var ssoDomain = sharedCookies.SsoDomain;
            if (!string.IsNullOrWhiteSpace(ssoDomain))
            {
                try
                {
                    var cookies = sharedCookies.Container.GetCookies(CreateSsoCookieUri(ssoDomain));
                    foreach (Cookie cookie in cookies)
                    {
                        cookie.Expired = true;
                    }
                }
                catch (UriFormatException)
                {
                    // Ignore invalid domain.
                }
            }

            cookieStore.Clear();
        }

        private static Uri CreateSsoCookieUri(string ssoDomain)
        {
            return new Uri($"https://{ssoDomain}");
        }
    }
}
