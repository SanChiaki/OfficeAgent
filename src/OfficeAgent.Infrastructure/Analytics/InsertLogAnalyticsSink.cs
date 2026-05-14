using System;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Infrastructure.Analytics
{
    public sealed class InsertLogAnalyticsSink : IAnalyticsSink
    {
        private static readonly JsonSerializerSettings AnalyticsJsonSettings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
        };

        private readonly Func<AppSettings> loadSettings;
        private readonly HttpClient httpClient;
        private readonly IAnalyticsProjectContextProvider projectContextProvider;

        public InsertLogAnalyticsSink(
            Func<AppSettings> loadSettings,
            HttpClient httpClient = null,
            CookieContainer cookieContainer = null,
            IAnalyticsProjectContextProvider projectContextProvider = null)
        {
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
            this.projectContextProvider = projectContextProvider;
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
                    Timeout = TimeSpan.FromSeconds(5),
                };
            }
            else
            {
                this.httpClient = new HttpClient
                {
                    Timeout = TimeSpan.FromSeconds(5),
                };
            }
        }

        public async Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
        {
            if (analyticsEvent == null)
            {
                throw new ArgumentNullException(nameof(analyticsEvent));
            }

            var settings = loadSettings() ?? new AppSettings();
            var analyticsUrl = AppSettings.NormalizeOptionalEndpointUrl(settings.AnalyticsUrl);
            if (!Uri.TryCreate(analyticsUrl, UriKind.Absolute, out var endpoint) ||
                (endpoint.Scheme != Uri.UriSchemeHttp && endpoint.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured Analytics URL is invalid. Update settings and try again.");
            }

            var payload = JsonConvert.SerializeObject(new
            {
                frontEndIntent = "excelAi",
                clientSource = "Excel",
                questionType = 1,
                askId = CreateRandomId(),
                talkId = CreateRandomId(),
                projectId = ResolveEnvelopeProjectId(analyticsEvent),
                answer = JsonConvert.SerializeObject(analyticsEvent, AnalyticsJsonSettings),
            });

            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                request.Content = new StringContent(payload, Encoding.UTF8, "application/json");

                using (var response = await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        var responseBody = await (response.Content?.ReadAsStringAsync() ?? Task.FromResult(string.Empty)).ConfigureAwait(false);
                        throw new InvalidOperationException(
                            $"Analytics request failed ({(int)response.StatusCode} {response.ReasonPhrase}): {responseBody}");
                    }
                }
            }
        }

        private static string CreateRandomId()
        {
            var bytes = new byte[24];
            using (var generator = RandomNumberGenerator.Create())
            {
                generator.GetBytes(bytes);
            }

            return Convert.ToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private string ResolveEnvelopeProjectId(AnalyticsEvent analyticsEvent)
        {
            var eventProjectId = ExtractProjectId(analyticsEvent?.Properties);
            if (!string.IsNullOrWhiteSpace(eventProjectId))
            {
                projectContextProvider?.RememberProjectId(eventProjectId);
                return eventProjectId;
            }

            var fallbackProjectId = projectContextProvider?.GetCurrentProjectId() ?? string.Empty;
            return string.IsNullOrWhiteSpace(fallbackProjectId) ? string.Empty : fallbackProjectId.Trim();
        }

        private static string ExtractProjectId(System.Collections.Generic.IDictionary<string, object> properties)
        {
            if (properties == null ||
                !properties.TryGetValue("projectId", out var projectIdValue) ||
                projectIdValue == null)
            {
                return string.Empty;
            }

            return projectIdValue.ToString()?.Trim() ?? string.Empty;
        }
    }
}
