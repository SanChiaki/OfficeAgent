using System;
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

        public InsertLogAnalyticsSink(Func<AppSettings> loadSettings, HttpClient httpClient = null)
        {
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
            this.httpClient = httpClient ?? new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(5),
            };
        }

        public async Task WriteAsync(AnalyticsEvent analyticsEvent, CancellationToken cancellationToken)
        {
            if (analyticsEvent == null)
            {
                throw new ArgumentNullException(nameof(analyticsEvent));
            }

            var settings = loadSettings() ?? new AppSettings();
            var baseUrl = AppSettings.NormalizeOptionalUrl(settings.AnalyticsBaseUrl);
            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured Analytics Base URL is invalid. Update settings and try again.");
            }

            var endpoint = new Uri($"{baseUri.AbsoluteUri.TrimEnd('/')}/insertLog");
            var payload = JsonConvert.SerializeObject(new
            {
                frontEndIntent = "excelAi",
                clientSource = "Excel",
                questionType = 1,
                askId = CreateRandomId(),
                talkId = CreateRandomId(),
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
    }
}
