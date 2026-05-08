using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Infrastructure.Http
{
    internal static class LlmApiFormat
    {
        public const string OpenAiCompatible = AppSettings.DefaultApiFormat;
        public const string AnthropicMessages = "anthropic-messages";
        public const string AnthropicVersion = "2023-06-01";
        public const int DefaultMaxTokens = 4096;

        public static bool IsAnthropicMessages(string apiFormat)
        {
            return string.Equals(AppSettings.NormalizeApiFormat(apiFormat), AnthropicMessages, StringComparison.Ordinal);
        }

        public static HttpRequestMessage CreateJsonRequest(Uri endpoint, string apiKey, string payload, string apiFormat)
        {
            var httpRequest = new HttpRequestMessage(HttpMethod.Post, endpoint);
            if (IsAnthropicMessages(apiFormat))
            {
                httpRequest.Headers.TryAddWithoutValidation("x-api-key", apiKey);
                httpRequest.Headers.TryAddWithoutValidation("anthropic-version", AnthropicVersion);
            }
            else
            {
                httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            }

            httpRequest.Content = new StringContent(payload, Encoding.UTF8, "application/json");
            return httpRequest;
        }

        public static Uri BuildChatCompletionsEndpoint(Uri baseUri)
        {
            var absoluteUri = baseUri.AbsoluteUri.TrimEnd('/');
            var absolutePath = baseUri.AbsolutePath?.Trim('/') ?? string.Empty;
            if (string.IsNullOrWhiteSpace(absolutePath))
            {
                return new Uri($"{absoluteUri}/v1/chat/completions");
            }

            return new Uri($"{absoluteUri}/chat/completions");
        }

        public static Uri BuildAnthropicMessagesEndpoint(Uri baseUri)
        {
            var absoluteUri = baseUri.AbsoluteUri.TrimEnd('/');
            var absolutePath = baseUri.AbsolutePath?.Trim('/') ?? string.Empty;
            if (string.IsNullOrWhiteSpace(absolutePath) ||
                (!string.Equals(absolutePath, "v1", StringComparison.OrdinalIgnoreCase) &&
                !absolutePath.EndsWith("/v1", StringComparison.OrdinalIgnoreCase)))
            {
                return new Uri($"{absoluteUri}/v1/messages");
            }

            return new Uri($"{absoluteUri}/messages");
        }

        public static string ExtractAnthropicMessageText(string responseBody, string errorPrefix)
        {
            try
            {
                var parsed = JObject.Parse(responseBody);
                var contentItems = parsed["content"] as JArray;
                if (contentItems == null)
                {
                    throw new InvalidOperationException($"{errorPrefix} returned an Anthropic message payload without content.");
                }

                foreach (var contentItem in contentItems)
                {
                    var contentType = contentItem["type"]?.Value<string>();
                    if (!string.Equals(contentType, "text", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var text = contentItem["text"]?.Value<string>();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text;
                    }
                }
            }
            catch (JsonException)
            {
                throw new InvalidOperationException($"{errorPrefix} returned a non-JSON Anthropic message payload.");
            }

            throw new InvalidOperationException($"{errorPrefix} returned an Anthropic message payload without text output.");
        }
    }
}
