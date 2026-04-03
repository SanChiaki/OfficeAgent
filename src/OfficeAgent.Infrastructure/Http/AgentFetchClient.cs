using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
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
                    SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                })
                {
                    Timeout = TimeSpan.FromSeconds(15),
                };
            }
            else
            {
                this.httpClient = new HttpClient(new HttpClientHandler
                {
                    SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                })
                {
                    Timeout = TimeSpan.FromSeconds(15),
                };
            }
        }

        public async Task<FetchResult> FetchAsync(string url, JObject headers = null)
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

                    ApplyCustomHeaders(request, headers);

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
                                ErrorMessage = $"请求失败：HTTP {(int)response.StatusCode} {response.ReasonPhrase}",
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
                    ErrorMessage = $"请求失败：请求超时（{httpClient.Timeout.TotalSeconds:0}秒）",
                };
            }
            catch (HttpRequestException error)
            {
                OfficeAgentLog.Error("agent_fetch", "request.exception", $"GET {uri} failed.", error);
                return new FetchResult
                {
                    Success = false,
                    ErrorMessage = $"请求失败：{error.Message}",
                };
            }
        }

        private static void ApplyCustomHeaders(HttpRequestMessage request, JObject headers)
        {
            if (headers == null || !headers.HasValues)
            {
                return;
            }

            foreach (var prop in headers.Properties())
            {
                var name = prop.Name;
                var value = prop.Value?.Value<string>() ?? string.Empty;

                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(value))
                {
                    continue;
                }

                // Skip headers that are managed by the HttpClient handler (cookies).
                if (string.Equals(name, "cookie", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, "set-cookie", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Handle headers that require specific HttpRequestMessage properties
                switch (name.ToLowerInvariant())
                {
                    case "authorization":
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", value);
                        break;
                    case "content-type":
                        if (request.Content != null)
                        {
                            request.Content.Headers.ContentType = new MediaTypeHeaderValue(value);
                        }
                        break;
                    default:
                        request.Headers.TryAddWithoutValidation(name, value);
                        break;
                }
            }
        }
    }
}
