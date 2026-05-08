using System;
using System.Net.Http;
using System.Security.Authentication;
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
            this.httpClient = httpClient ?? new HttpClient(new HttpClientHandler
            {
                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
            })
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
                if (LlmApiFormat.IsAnthropicMessages(settings.ApiFormat))
                {
                    return await CompleteWithAnthropicMessagesAsync(baseUri, settings, request).ConfigureAwait(false);
                }

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

            var responseBody = await SendRequestAsync(
                endpoint,
                settings.ApiKey,
                payload,
                allowLegacyFallback: true,
                apiFormat: settings.ApiFormat).ConfigureAwait(false);
            return ExtractChatCompletionsText(responseBody);
        }

        private async Task<string> CompleteWithAnthropicMessagesAsync(Uri baseUri, AppSettings settings, PlannerRequest request)
        {
            var endpoint = LlmApiFormat.BuildAnthropicMessagesEndpoint(baseUri);
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                max_tokens = LlmApiFormat.DefaultMaxTokens,
                system = BuildPlannerInstructions(),
                messages = BuildAnthropicMessages(request),
            });

            var responseBody = await SendRequestAsync(
                endpoint,
                settings.ApiKey,
                payload,
                allowLegacyFallback: false,
                apiFormat: settings.ApiFormat).ConfigureAwait(false);
            return LlmApiFormat.ExtractAnthropicMessageText(responseBody, "Planner API");
        }

        private async Task<string> CompleteWithLegacyPlannerAsync(Uri baseUri, AppSettings settings, PlannerRequest request)
        {
            var endpoint = new Uri($"{baseUri.AbsoluteUri.TrimEnd('/')}/planner");
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                request,
            });
            return await SendRequestAsync(
                endpoint,
                settings.ApiKey,
                payload,
                allowLegacyFallback: false,
                apiFormat: settings.ApiFormat).ConfigureAwait(false);
        }

        private async Task<string> SendRequestAsync(Uri endpoint, string apiKey, string payload, bool allowLegacyFallback, string apiFormat)
        {
            using (var httpRequest = LlmApiFormat.CreateJsonRequest(endpoint, apiKey, payload, apiFormat))
            {
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
            return LlmApiFormat.BuildChatCompletionsEndpoint(baseUri);
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

        private static object[] BuildAnthropicMessages(PlannerRequest request)
        {
            var turns = new System.Collections.Generic.List<Tuple<string, string>>();
            var hasSeenUserTurn = false;
            foreach (var turn in request.ConversationHistory ?? System.Array.Empty<ConversationTurn>())
            {
                if (string.IsNullOrWhiteSpace(turn.Role) || string.IsNullOrWhiteSpace(turn.Content))
                {
                    continue;
                }

                var role = string.Equals(turn.Role, "assistant", StringComparison.OrdinalIgnoreCase) ? "assistant" : "user";
                if (!hasSeenUserTurn)
                {
                    if (!string.Equals(role, "user", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    hasSeenUserTurn = true;
                }

                AddAnthropicTurn(turns, role, turn.Content);
            }

            AddAnthropicTurn(turns, "user", BuildPlannerPrompt(request));

            var messages = new System.Collections.Generic.List<object>();
            foreach (var turn in turns)
            {
                messages.Add(CreateChatMessage(turn.Item1, turn.Item2));
            }

            return messages.ToArray();
        }

        private static void AddAnthropicTurn(System.Collections.Generic.List<Tuple<string, string>> turns, string role, string content)
        {
            if (turns.Count > 0 && string.Equals(turns[turns.Count - 1].Item1, role, StringComparison.Ordinal))
            {
                var last = turns[turns.Count - 1];
                turns[turns.Count - 1] = Tuple.Create(last.Item1, $"{last.Item2}\n\n{content}");
                return;
            }

            turns.Add(Tuple.Create(role, content));
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
                + "When the user explicitly asks to include custom request headers (e.g., 'with header X-Api-Key: abc'), add a headers object to fetch.url args: {\"type\":\"fetch.url\",\"args\":{\"url\":\"http://...\",\"headers\":{\"X-Api-Key\":\"abc\"}}}. Headers is an optional flat object of string key-value pairs, only present when the user specifies headers. "
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
