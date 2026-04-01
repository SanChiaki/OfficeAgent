using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class LlmPlannerClient : ILlmPlannerClient
    {
        private readonly HttpClient httpClient;
        private readonly Func<AppSettings> loadSettings;

        public LlmPlannerClient(FileSettingsStore settingsStore, HttpClient httpClient = null)
            : this(httpClient, () => settingsStore?.Load() ?? new AppSettings())
        {
        }

        public LlmPlannerClient(HttpClient httpClient, Func<AppSettings> loadSettings)
        {
            this.httpClient = httpClient ?? new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(120),
            };
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
        }

        public string Complete(PlannerRequest request)
        {
            return CompleteAsync(request).GetAwaiter().GetResult();
        }

        public async Task<string> CompleteAsync(PlannerRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var settings = loadSettings() ?? new AppSettings();
            var baseUrl = AppSettings.NormalizeBaseUrl(settings.BaseUrl);
            if (string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                throw new InvalidOperationException("An API Key is required before agent planning can call the planner API.");
            }

            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured Planner API Base URL is invalid. Update settings and try again.");
            }

            try
            {
                return await CompleteWithOpenAiCompatibleChatCompletionsAsync(baseUri, settings, request).ConfigureAwait(false);
            }
            catch (LegacyPlannerFallbackException)
            {
                return await CompleteWithLegacyPlannerAsync(baseUri, settings, request).ConfigureAwait(false);
            }
        }

        private async Task<string> CompleteWithOpenAiCompatibleChatCompletionsAsync(Uri baseUri, AppSettings settings, PlannerRequest request)
        {
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

            var responseBody = await SendRequestAsync(endpoint, settings.ApiKey, payload, allowLegacyFallback: true).ConfigureAwait(false);
            return ExtractChatCompletionsText(responseBody);
        }

        private async Task<string> CompleteWithLegacyPlannerAsync(Uri baseUri, AppSettings settings, PlannerRequest request)
        {
            var endpoint = new Uri($"{baseUri.AbsoluteUri.TrimEnd('/')}/planner");
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                request,
            });
            return await SendRequestAsync(endpoint, settings.ApiKey, payload, allowLegacyFallback: false).ConfigureAwait(false);
        }

        private async Task<string> SendRequestAsync(Uri endpoint, string apiKey, string payload, bool allowLegacyFallback)
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
                        if (allowLegacyFallback &&
                            ((int)response.StatusCode == 404 || (int)response.StatusCode == 405))
                        {
                            throw new LegacyPlannerFallbackException();
                        }

                        throw new InvalidOperationException(
                            $"Planner API request failed ({(int)response.StatusCode} {response.ReasonPhrase}): {responseBody}");
                    }

                    if (string.IsNullOrWhiteSpace(responseBody))
                    {
                        throw new InvalidOperationException("Planner API returned an empty response body.");
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

        private static object CreateChatMessage(string role, string text)
        {
            return new
            {
                role,
                content = text,
            };
        }

        private static string BuildPlannerPrompt(PlannerRequest request)
        {
            return "Planner request:\n" + JsonConvert.SerializeObject(
                request ?? new PlannerRequest(),
                Formatting.Indented,
                new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                });
        }

        private static object[] BuildChatMessages(PlannerRequest request)
        {
            var messages = new System.Collections.Generic.List<object>();
            messages.Add(CreateChatMessage("system", BuildPlannerInstructions()));

            foreach (var turn in request.ConversationHistory ?? System.Array.Empty<ConversationTurn>())
            {
                if (!string.IsNullOrWhiteSpace(turn.Role) && !string.IsNullOrWhiteSpace(turn.Content))
                {
                    messages.Add(CreateChatMessage(turn.Role, turn.Content));
                }
            }

            messages.Add(CreateChatMessage("user", BuildPlannerPrompt(request)));
            return messages.ToArray();
        }

        private static string BuildPlannerInstructions()
        {
            return "You are OfficeAgent's planner. "
                + "Return exactly one JSON object and no markdown. "
                + "Always include the keys mode, assistantMessage, step, and plan. "
                + "Use null for step or plan when they do not apply. "
                + "assistantMessage should be concise and use the user's language when possible. "
                + "Supported modes are message, read_step, and plan. "
                + "Use message when no Excel action is needed or the request is unsupported. "
                + "Use read_step when you need data before planning. "
                + "Supported read_step types are: "
                + "1. excel.readSelectionTable with empty args — reads the user's current Excel selection as a table. "
                + "2. excel.readRange with args { address: string, sheetName?: string } — reads a specific range from any worksheet. sheetName defaults to the active sheet if omitted. "
                + "3. fetch.url with args { url: string } — makes an HTTP GET request to fetch external data. The url must be a full absolute URL (e.g. \"http://localhost:3200/api/performance\"). "
                + "Use plan for any write or side-effect sequence. "
                + "Supported plan step types are excel.writeRange, excel.addWorksheet, excel.renameWorksheet, excel.deleteWorksheet, and skill.upload_data. "
                + "Never invent other step types. "
                + "For excel.writeRange use args targetAddress and values. "
                + "For excel.addWorksheet use arg newSheetName. "
                + "For excel.renameWorksheet use args sheetName and newSheetName. "
                + "For excel.deleteWorksheet use arg sheetName. "
                + "For skill.upload_data use arg userInput and preserve the user's upload intent. "
                + "Use the provided selection metadata, headers, sample rows, prior observations, and apiBaseUrl. "
                + "Only request read_step when the summary is insufficient. "
                + "When mode=read_step, set step to {\"type\":\"excel.readSelectionTable\",\"args\":{}} or {\"type\":\"excel.readRange\",\"args\":{\"address\":\"A1:D10\"}} or {\"type\":\"fetch.url\",\"args\":{\"url\":\"http://...\"}} and set plan to null. "
                + "When mode=plan, set plan.summary and plan.steps, and set step to null. "
                + "When mode=message, set both step and plan to null. "
                + "If the request cannot be completed safely with the supported actions, answer with mode=message. "
                + "Prior conversation turns are included as context. Use them to understand follow-up questions and maintain coherence. "
                + "The apiBaseUrl field in the request indicates the configured business API endpoint. Prefer using it as the base for fetch.url requests when the user asks about business data.";
        }

        private static string ExtractChatCompletionsText(string responseBody)
        {
            try
            {
                var parsed = JObject.Parse(responseBody);
                var content = parsed["choices"]?[0]?["message"]?["content"];
                if (content == null)
                {
                    throw new InvalidOperationException("Planner API returned a chat completion payload without message content.");
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
                throw new InvalidOperationException("Planner API returned a non-JSON chat completion payload.");
            }

            throw new InvalidOperationException("Planner API returned a chat completion payload without planner text output.");
        }

        private sealed class LegacyPlannerFallbackException : Exception
        {
        }
    }
}
