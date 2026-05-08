using System;
using System.IO;
using System.Net.Http;
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

            var apiFormat = AppSettings.NormalizeApiFormat(settings.ApiFormat);
            var endpoint = LlmApiFormat.IsAnthropicMessages(apiFormat)
                ? LlmApiFormat.BuildAnthropicMessagesEndpoint(baseUri)
                : LlmApiFormat.BuildChatCompletionsEndpoint(baseUri);
            var mappingJson = await TrySendStreamingRequestAsync(endpoint, settings.ApiKey, settings.Model, request, apiFormat).ConfigureAwait(false);
            if (mappingJson == null)
            {
                var responseBody = await SendRequestAsync(
                    endpoint,
                    settings.ApiKey,
                    BuildPayload(settings.Model, request, stream: false, apiFormat),
                    apiFormat).ConfigureAwait(false);
                mappingJson = LlmApiFormat.IsAnthropicMessages(apiFormat)
                    ? LlmApiFormat.ExtractAnthropicMessageText(responseBody, "AI column mapping API")
                    : ExtractChatCompletionsText(responseBody);
            }

            return ParseMappingResponse(mappingJson);
        }

        private async Task<string> TrySendStreamingRequestAsync(Uri endpoint, string apiKey, string model, AiColumnMappingRequest request, string apiFormat)
        {
            using (var httpRequest = CreateRequest(endpoint, apiKey, BuildPayload(model, request, stream: true, apiFormat), apiFormat))
            {
                using (var response = await httpClient.SendAsync(httpRequest, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        if (ShouldFallbackToNonStreaming(response.StatusCode))
                        {
                            return null;
                        }

                        var responseBody = await (response.Content?.ReadAsStringAsync() ?? Task.FromResult(string.Empty)).ConfigureAwait(false);
                        throw FormatRequestFailure(response, responseBody);
                    }

                    return await ReadStreamingMappingTextAsync(response, apiFormat).ConfigureAwait(false);
                }
            }
        }

        private async Task<string> SendRequestAsync(Uri endpoint, string apiKey, string payload, string apiFormat)
        {
            using (var httpRequest = CreateRequest(endpoint, apiKey, payload, apiFormat))
            {
                using (var response = await httpClient.SendAsync(httpRequest).ConfigureAwait(false))
                {
                    var responseBody = await (response.Content?.ReadAsStringAsync() ?? Task.FromResult(string.Empty)).ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw FormatRequestFailure(response, responseBody);
                    }

                    if (string.IsNullOrWhiteSpace(responseBody))
                    {
                        throw new InvalidOperationException("AI column mapping API returned an empty response body.");
                    }

                    return responseBody;
                }
            }
        }

        private static HttpRequestMessage CreateRequest(Uri endpoint, string apiKey, string payload, string apiFormat)
        {
            return LlmApiFormat.CreateJsonRequest(endpoint, apiKey, payload, apiFormat);
        }

        private static string BuildPayload(string model, AiColumnMappingRequest request, bool stream, string apiFormat)
        {
            if (LlmApiFormat.IsAnthropicMessages(apiFormat))
            {
                return JsonConvert.SerializeObject(new
                {
                    model,
                    max_tokens = LlmApiFormat.DefaultMaxTokens,
                    system = BuildInstructions(),
                    messages = new[]
                    {
                        CreateChatMessage("user", BuildMappingPrompt(request)),
                    },
                    stream,
                });
            }

            return JsonConvert.SerializeObject(new
            {
                model,
                messages = BuildChatMessages(request),
                response_format = new
                {
                    type = "json_object",
                },
                stream,
            });
        }

        private static bool ShouldFallbackToNonStreaming(System.Net.HttpStatusCode statusCode)
        {
            return statusCode == System.Net.HttpStatusCode.BadRequest ||
                   statusCode == System.Net.HttpStatusCode.NotFound ||
                   statusCode == System.Net.HttpStatusCode.NotAcceptable ||
                   statusCode == System.Net.HttpStatusCode.UnsupportedMediaType ||
                   (int)statusCode == 422;
        }

        private static InvalidOperationException FormatRequestFailure(HttpResponseMessage response, string responseBody)
        {
            return new InvalidOperationException(
                $"AI column mapping API request failed ({(int)response.StatusCode} {response.ReasonPhrase}): {responseBody}");
        }

        private static async Task<string> ReadStreamingMappingTextAsync(HttpResponseMessage response, string apiFormat)
        {
            var builder = new StringBuilder();
            var rawBody = new StringBuilder();
            using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync().ConfigureAwait(false);
                    if (line == null)
                    {
                        break;
                    }

                    rawBody.AppendLine(line);
                    var trimmed = line.Trim();
                    if (trimmed.Length == 0 || trimmed.StartsWith(":", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (!trimmed.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var data = trimmed.Substring("data:".Length).Trim();
                    if (string.Equals(data, "[DONE]", StringComparison.Ordinal))
                    {
                        break;
                    }

                    var isComplete = LlmApiFormat.IsAnthropicMessages(apiFormat)
                        ? AppendAnthropicStreamingDeltaContent(builder, data)
                        : AppendStreamingDeltaContent(builder, data);
                    if (isComplete)
                    {
                        break;
                    }
                }
            }

            if (builder.Length == 0)
            {
                var fallbackBody = rawBody.ToString();
                if (!string.IsNullOrWhiteSpace(fallbackBody))
                {
                    return LlmApiFormat.IsAnthropicMessages(apiFormat)
                        ? LlmApiFormat.ExtractAnthropicMessageText(fallbackBody, "AI column mapping API")
                        : ExtractChatCompletionsText(fallbackBody);
                }

                throw new InvalidOperationException("AI column mapping API returned an empty streaming response.");
            }

            return builder.ToString();
        }

        private static bool AppendStreamingDeltaContent(StringBuilder builder, string data)
        {
            try
            {
                var parsed = JObject.Parse(data);
                var choice = parsed["choices"]?.First;
                if (choice == null)
                {
                    return false;
                }

                var content = choice["delta"]?["content"] ?? choice["message"]?["content"];
                if (content == null)
                {
                    return IsStreamingChoiceFinished(choice);
                }

                if (content.Type == JTokenType.String)
                {
                    builder.Append(content.Value<string>());
                    return IsStreamingChoiceFinished(choice);
                }

                if (content is JArray contentItems)
                {
                    foreach (var contentItem in contentItems)
                    {
                        var contentText = contentItem["text"]?.Value<string>();
                        if (!string.IsNullOrEmpty(contentText))
                        {
                            builder.Append(contentText);
                        }
                    }
                }

                return IsStreamingChoiceFinished(choice);
            }
            catch (JsonException error)
            {
                throw new InvalidOperationException("AI column mapping API returned a malformed streaming chunk.", error);
            }
        }

        private static bool AppendAnthropicStreamingDeltaContent(StringBuilder builder, string data)
        {
            try
            {
                var parsed = JObject.Parse(data);
                var eventType = parsed["type"]?.Value<string>();
                if (string.Equals(eventType, "message_stop", StringComparison.Ordinal))
                {
                    return true;
                }

                if (string.Equals(eventType, "error", StringComparison.Ordinal))
                {
                    var message = parsed["error"]?["message"]?.Value<string>() ?? data;
                    throw new InvalidOperationException($"AI column mapping API returned a streaming error: {message}");
                }

                var delta = parsed["delta"];
                var deltaType = delta?["type"]?.Value<string>();
                if (string.Equals(eventType, "content_block_delta", StringComparison.Ordinal) &&
                    string.Equals(deltaType, "text_delta", StringComparison.Ordinal))
                {
                    var text = delta["text"]?.Value<string>();
                    if (!string.IsNullOrEmpty(text))
                    {
                        builder.Append(text);
                    }
                }

                return string.Equals(parsed["stop_reason"]?.Value<string>(), "end_turn", StringComparison.Ordinal);
            }
            catch (JsonException error)
            {
                throw new InvalidOperationException("AI column mapping API returned a malformed Anthropic streaming chunk.", error);
            }
        }

        private static bool IsStreamingChoiceFinished(JToken choice)
        {
            var finishReason = choice?["finish_reason"]?.Value<string>();
            return !string.IsNullOrWhiteSpace(finishReason);
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
                Formatting.None,
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
