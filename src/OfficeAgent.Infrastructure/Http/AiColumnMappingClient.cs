using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class AiColumnMappingClient : IAiColumnMappingClient
    {
        private readonly HttpClient httpClient;
        private readonly Func<AppSettings> loadSettings;

        public AiColumnMappingClient(FileSettingsStore settingsStore, HttpClient httpClient = null)
            : this(httpClient, () => settingsStore?.Load() ?? new AppSettings())
        {
        }

        public AiColumnMappingClient(Func<AppSettings> loadSettings, HttpClient httpClient = null)
            : this(httpClient, loadSettings)
        {
        }

        public AiColumnMappingClient(HttpClient httpClient, Func<AppSettings> loadSettings)
        {
            this.httpClient = httpClient ?? new HttpClient(new HttpClientHandler
            {
                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
            })
            {
                Timeout = TimeSpan.FromSeconds(120),
            };
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
        }

        public AiColumnMappingResponse Map(AiColumnMappingRequest request)
        {
            return MapAsync(request).GetAwaiter().GetResult();
        }

        public async Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var settings = loadSettings() ?? new AppSettings();
            var baseUrl = AppSettings.NormalizeBaseUrl(settings.BaseUrl);
            if (string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                throw new InvalidOperationException("An API Key is required before AI column mapping can call the mapping API.");
            }

            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured AI column mapping API Base URL is invalid. Update settings and try again.");
            }

            var endpoint = BuildChatCompletionsEndpoint(baseUri);
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                messages = BuildChatMessages(request),
                response_format = new
                {
                    type = "json_object",
                },
            });
            var responseBody = await SendRequestAsync(endpoint, settings.ApiKey, payload).ConfigureAwait(false);
            var mappingJson = ExtractChatCompletionsText(responseBody);
            return ParseMappingResponse(mappingJson);
        }

        private async Task<string> SendRequestAsync(Uri endpoint, string apiKey, string payload)
        {
            using (var httpRequest = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                httpRequest.Content = new StringContent(payload, Encoding.UTF8, "application/json");

                using (var response = await httpClient.SendAsync(httpRequest).ConfigureAwait(false))
                {
                    var responseBody = await (response.Content?.ReadAsStringAsync() ?? Task.FromResult(string.Empty)).ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            $"AI column mapping API request failed ({(int)response.StatusCode} {response.ReasonPhrase}): {responseBody}");
                    }

                    if (string.IsNullOrWhiteSpace(responseBody))
                    {
                        throw new InvalidOperationException("AI column mapping API returned an empty response body.");
                    }

                    return responseBody;
                }
            }
        }

        private static Uri BuildChatCompletionsEndpoint(Uri baseUri)
        {
            var absoluteUri = baseUri.AbsoluteUri.TrimEnd('/');
            var absolutePath = baseUri.AbsolutePath?.Trim('/') ?? string.Empty;
            if (string.IsNullOrWhiteSpace(absolutePath))
            {
                return new Uri($"{absoluteUri}/v1/chat/completions");
            }

            return new Uri($"{absoluteUri}/chat/completions");
        }

        private static object[] BuildChatMessages(AiColumnMappingRequest request)
        {
            return new[]
            {
                CreateChatMessage("system", BuildInstructions()),
                CreateChatMessage("user", BuildMappingPrompt(request)),
            };
        }

        private static object CreateChatMessage(string role, string text)
        {
            return new
            {
                role,
                content = text,
            };
        }

        private static string BuildMappingPrompt(AiColumnMappingRequest request)
        {
            return "Column mapping request:\n" + JsonConvert.SerializeObject(
                request ?? new AiColumnMappingRequest(),
                Formatting.Indented,
                new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                });
        }

        private static string BuildInstructions()
        {
            return "You map actual Excel headers to known OfficeAgent SheetFieldMappings candidates. "
                + "Return exactly one JSON object and no markdown. "
                + "The response object must include Mappings and Unmatched arrays. "
                + "Each mapping must include ExcelColumn, ActualL1, ActualL2, TargetHeaderId, TargetApiFieldKey, Confidence, and Reason. "
                + "Use the provided Candidate HeaderId and ApiFieldKey exactly; never invent target identities. "
                + "Confidence is a number between 0 and 1. "
                + "Only map a header when the business meaning is clear. Put uncertain headers in Unmatched. "
                + "For two-row headers, preserve the actual parent text as ActualL1 and child text as ActualL2.";
        }

        private static string ExtractChatCompletionsText(string responseBody)
        {
            try
            {
                var parsed = JObject.Parse(responseBody);
                var content = parsed["choices"]?[0]?["message"]?["content"];
                if (content == null)
                {
                    throw new InvalidOperationException("AI column mapping API returned a chat completion payload without message content.");
                }

                if (content.Type == JTokenType.String)
                {
                    return content.Value<string>();
                }

                if (content is JArray contentItems)
                {
                    foreach (var contentItem in contentItems)
                    {
                        var contentType = contentItem["type"]?.Value<string>();
                        if (!string.Equals(contentType, "text", StringComparison.Ordinal) &&
                            !string.Equals(contentType, "output_text", StringComparison.Ordinal))
                        {
                            continue;
                        }

                        var contentText = contentItem["text"]?.Value<string>();
                        if (!string.IsNullOrWhiteSpace(contentText))
                        {
                            return contentText;
                        }
                    }
                }
            }
            catch (JsonException)
            {
                throw new InvalidOperationException("AI column mapping API returned a non-JSON chat completion payload.");
            }

            throw new InvalidOperationException("AI column mapping API returned a chat completion payload without mapping text output.");
        }

        private static AiColumnMappingResponse ParseMappingResponse(string mappingJson)
        {
            try
            {
                var parsed = JObject.Parse(mappingJson);
                if (parsed["Mappings"] == null)
                {
                    throw new InvalidOperationException("AI column mapping API returned a mapping payload without Mappings.");
                }

                var response = parsed.ToObject<AiColumnMappingResponse>();
                if (response == null || response.Mappings == null)
                {
                    throw new InvalidOperationException("AI column mapping API returned an invalid mapping payload.");
                }

                return new AiColumnMappingResponse
                {
                    Mappings = response.Mappings ?? Array.Empty<AiColumnMappingSuggestion>(),
                    Unmatched = response.Unmatched ?? Array.Empty<AiColumnMappingUnmatchedHeader>(),
                };
            }
            catch (JsonException error)
            {
                throw new InvalidOperationException("AI column mapping API returned malformed mapping JSON.", error);
            }
        }
    }
}
