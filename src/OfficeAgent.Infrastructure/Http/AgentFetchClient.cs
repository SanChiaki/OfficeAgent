using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class AgentFetchClient : IAgentFetchClient
    {
        private readonly HttpClient httpClient;
        private readonly Func<AppSettings> loadSettings;

        public AgentFetchClient(Func<AppSettings> loadSettings, CookieContainer cookieContainer = null, HttpClient httpClient = null)
        {
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));

            if (httpClient != null)
            {
                this.httpClient = httpClient;
            }
            else if (cookieContainer != null)
            {
                this.httpClient = new HttpClient(new HttpClientHandler
                {
                    CookieContainer = cookieContainer,
                    UseCookies = true,
                })
                {
                    Timeout = TimeSpan.FromSeconds(15),
                };
            }
            else
            {
                this.httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
            }
        }

        public async Task<FetchResult> FetchAsync(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return new FetchResult
                {
                    Success = false,
                    ErrorMessage = "URL is required.",
                };
            }

            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                return new FetchResult
                {
                    Success = false,
                    ErrorMessage = $"Invalid URL: {url}",
                };
            }

            var settings = loadSettings() ?? new AppSettings();

            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, uri))
                {
                    if (!string.IsNullOrWhiteSpace(settings.ApiKey))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                    }

                    OfficeAgentLog.Info("agent_fetch", "request.begin", $"GET {uri}");

                    using (var response = await httpClient.SendAsync(request).ConfigureAwait(false))
                    {
                        var body = await (response.Content?.ReadAsStringAsync() ?? Task.FromResult(string.Empty)).ConfigureAwait(false);

                        OfficeAgentLog.Info("agent_fetch", "request.completed", $"GET {uri} — {(int)response.StatusCode}");

                        if (!response.IsSuccessStatusCode)
                        {
                            return new FetchResult
                            {
                                Success = false,
                                StatusCode = (int)response.StatusCode,
                                Body = body,
                                ErrorMessage = $"HTTP {(int)response.StatusCode} {response.ReasonPhrase}",
                            };
                        }

                        return new FetchResult
                        {
                            Success = true,
                            StatusCode = (int)response.StatusCode,
                            Body = body,
                        };
                    }
                }
            }
            catch (TaskCanceledException error)
            {
                OfficeAgentLog.Error("agent_fetch", "request.timeout", $"GET {uri} timed out.", error);
                return new FetchResult
                {
                    Success = false,
                    ErrorMessage = $"Request timed out after {httpClient.Timeout.TotalSeconds:0} seconds.",
                };
            }
            catch (HttpRequestException error)
            {
                OfficeAgentLog.Error("agent_fetch", "request.exception", $"GET {uri} failed.", error);
                return new FetchResult
                {
                    Success = false,
                    ErrorMessage = $"Request failed: {error.Message}",
                };
            }
        }
    }
}
