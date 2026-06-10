using System;
using System.IO;
using System.Net;
using OfficeAgent.Infrastructure.Http;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class AccountSessionServiceTests : IDisposable
    {
        private readonly string tempDirectory;

        public AccountSessionServiceTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.AccountSession.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDirectory);
        }

        [Fact]
        public void LogoutExpiresSharedCookiesAndClearsPersistentCookieFile()
        {
            var sharedCookies = new SharedCookieContainer { SsoDomain = "example.com" };
            sharedCookies.Container.Add(
                new Uri("https://example.com"),
                new Cookie("sso-token", "token-123", "/"));
            var cookieStorePath = Path.Combine(tempDirectory, "cookies.json");
            File.WriteAllText(cookieStorePath, "persisted-cookie-data");
            var cookieStore = new FileCookieStore(cookieStorePath, new DpapiSecretProtector());
            var sessionService = new AccountSessionService(sharedCookies, cookieStore);

            Assert.True(sessionService.IsLoggedIn());

            sessionService.Logout();

            Assert.False(sessionService.IsLoggedIn());
            Assert.Empty(sharedCookies.Container.GetCookies(new Uri("https://example.com")));
            Assert.False(File.Exists(cookieStorePath));
            Assert.False(sessionService.IsServerAuthenticated);
        }

        [Fact]
        public void ServerAuthenticationStateDoesNotUseCookiePresenceAsLoginTruth()
        {
            var sharedCookies = new SharedCookieContainer { SsoDomain = "example.com" };
            sharedCookies.Container.Add(
                new Uri("https://example.com"),
                new Cookie("sso-token", "token-123", "/"));
            var cookieStore = new FileCookieStore(Path.Combine(tempDirectory, "cookies.json"), new DpapiSecretProtector());
            var sessionService = new AccountSessionService(sharedCookies, cookieStore);

            Assert.True(sessionService.IsLoggedIn());
            Assert.False(sessionService.IsServerAuthenticated);

            sessionService.MarkServerAuthenticated();

            Assert.True(sessionService.IsServerAuthenticated);

            sessionService.MarkServerAuthenticationRequired();

            Assert.False(sessionService.IsServerAuthenticated);
            Assert.False(sessionService.IsLoggedIn());
            Assert.Empty(sharedCookies.Container.GetCookies(new Uri("https://example.com")));
        }

        [Theory]
        [InlineData("")]
        [InlineData("not a uri")]
        public void ConfigureSsoDomainClearsInvalidSsoUrls(string ssoUrl)
        {
            var sharedCookies = new SharedCookieContainer { SsoDomain = "example.com" };
            var cookieStore = new FileCookieStore(Path.Combine(tempDirectory, "cookies.json"), new DpapiSecretProtector());
            var sessionService = new AccountSessionService(sharedCookies, cookieStore);

            sessionService.ConfigureSsoDomain(ssoUrl);

            Assert.Equal(string.Empty, sharedCookies.SsoDomain);
            Assert.False(sessionService.IsLoggedIn());
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }
    }
}
