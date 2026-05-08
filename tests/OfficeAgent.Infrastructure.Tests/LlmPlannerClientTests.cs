using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class LlmPlannerClientTests
    {
        [Fact]
        public void CompletePostsPlannerRequestsToTheConfiguredEndpoint()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"id\":\"chatcmpl-123\","
                    + "\"object\":\"chat.completion\","
                    + "\"created\":1739177682,"
                    + "\"model\":\"gpt-5-mini\","
                    + "\"choices\":[{"
                    + "\"index\":0,"
                    + "\"message\":{"
                    + "\"role\":\"assistant\","
                    + "\"content\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\",\\\"step\\\":null,\\\"plan\\\":null}\""
                    + "},"
                    + "\"finish_reason\":\"stop\""
                    + "}]"
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = " https://api.internal.example/ ",
                    Model = "gpt-5-mini",
                });

            var response = client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal("https://api.internal.example/v1/chat/completions", handler.LastRequest.RequestUri.ToString());
            Assert.Equal("Bearer", handler.LastRequest.Headers.Authorization?.Scheme);
            Assert.Equal("secret-token", handler.LastRequest.Headers.Authorization?.Parameter);
            Assert.Contains("Create a summary sheet", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("gpt-5-mini", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"response_format\":{\"type\":\"json_object\"}", handler.LastBody, StringComparison.Ordinal);
            Assert.DoesNotContain("json_schema", handler.LastBody, StringComparison.Ordinal);
            Assert.Equal("{\"mode\":\"message\",\"assistantMessage\":\"ok\",\"step\":null,\"plan\":null}", response);
        }

        [Fact]
        public void CompletePreservesBaseUrlPathPrefixes()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"id\":\"chatcmpl-123\","
                    + "\"object\":\"chat.completion\","
                    + "\"created\":1739177682,"
                    + "\"model\":\"gpt-5-mini\","
                    + "\"choices\":[{"
                    + "\"index\":0,"
                    + "\"message\":{"
                    + "\"role\":\"assistant\","
                    + "\"content\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\",\\\"step\\\":null,\\\"plan\\\":null}\""
                    + "},"
                    + "\"finish_reason\":\"stop\""
                    + "}]"
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example/v1/",
                    Model = "gpt-5-mini",
                });

            client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal("https://api.internal.example/v1/chat/completions", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void CompleteFallsBackToTheLegacyPlannerEndpointWhenChatCompletionsApiIsUnavailable()
        {
            var handler = new RecordingHandler(request =>
            {
                if (request.RequestUri.ToString() == "https://api.internal.example/v1/chat/completions")
                {
                    return new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent("{\"error\":{\"message\":\"not found\"}}"),
                    };
                }

                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}"),
                };
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var response = client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal(2, handler.CallCount);
            Assert.Equal("https://api.internal.example/planner", handler.LastRequest.RequestUri.ToString());
            Assert.Equal("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}", response);
        }

        [Fact]
        public void CompletePostsAnthropicMessagesRequestsWhenConfigured()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"id\":\"msg_123\","
                    + "\"type\":\"message\","
                    + "\"role\":\"assistant\","
                    + "\"content\":[{\"type\":\"text\",\"text\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\",\\\"step\\\":null,\\\"plan\\\":null}\"}],"
                    + "\"stop_reason\":\"end_turn\""
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = " https://api.anthropic.com/ ",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            var response = client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal("https://api.anthropic.com/v1/messages", handler.LastRequest.RequestUri.ToString());
            Assert.Null(handler.LastRequest.Headers.Authorization);
            Assert.True(handler.LastRequest.Headers.TryGetValues("x-api-key", out var apiKeyValues));
            Assert.Contains("secret-token", apiKeyValues);
            Assert.True(handler.LastRequest.Headers.TryGetValues("anthropic-version", out var versionValues));
            Assert.Contains("2023-06-01", versionValues);
            Assert.Contains("\"system\":", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"messages\":[", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"max_tokens\":4096", handler.LastBody, StringComparison.Ordinal);
            Assert.DoesNotContain("response_format", handler.LastBody, StringComparison.Ordinal);
            Assert.Equal("{\"mode\":\"message\",\"assistantMessage\":\"ok\",\"step\":null,\"plan\":null}", response);
        }

        [Fact]
        public void CompleteBuildsAnthropicMessagesFromUserFirstConversationHistory()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"id\":\"msg_123\","
                    + "\"type\":\"message\","
                    + "\"role\":\"assistant\","
                    + "\"content\":[{\"type\":\"text\",\"text\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\",\\\"step\\\":null,\\\"plan\\\":null}\"}],"
                    + "\"stop_reason\":\"end_turn\""
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.anthropic.com",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
                ConversationHistory = new[]
                {
                    new ConversationTurn { Role = "assistant", Content = "Welcome" },
                    new ConversationTurn { Role = "user", Content = "First request" },
                    new ConversationTurn { Role = "user", Content = "Extra context" },
                    new ConversationTurn { Role = "assistant", Content = "Assistant answer" },
                },
            });

            Assert.DoesNotContain("\"role\":\"assistant\",\"content\":\"Welcome\"", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"role\":\"user\",\"content\":\"First request\\n\\nExtra context\"", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"role\":\"assistant\",\"content\":\"Assistant answer\"", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"role\":\"user\",\"content\":\"Planner request:", handler.LastBody, StringComparison.Ordinal);
        }

        [Fact]
        public void CompleteAppendsAnthropicVersionPathAfterProxyPrefixes()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"id\":\"msg_123\","
                    + "\"type\":\"message\","
                    + "\"role\":\"assistant\","
                    + "\"content\":[{\"type\":\"text\",\"text\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\",\\\"step\\\":null,\\\"plan\\\":null}\"}],"
                    + "\"stop_reason\":\"end_turn\""
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example/anthropic/",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal("https://api.internal.example/anthropic/v1/messages", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void CompleteDoesNotDuplicateAnthropicVersionPath()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"id\":\"msg_123\","
                    + "\"type\":\"message\","
                    + "\"role\":\"assistant\","
                    + "\"content\":[{\"type\":\"text\",\"text\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\",\\\"step\\\":null,\\\"plan\\\":null}\"}],"
                    + "\"stop_reason\":\"end_turn\""
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.anthropic.com/v1/",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal("https://api.anthropic.com/v1/messages", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void CompleteRejectsMissingApiKeys()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = " ",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            Action action = () => client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            var error = Assert.Throws<InvalidOperationException>(action);

            Assert.Contains("API Key", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public void CompleteRejectsInvalidBaseUrls()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "api.internal.example",
                    Model = "gpt-5-mini",
                });

            Action action = () => client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            var error = Assert.Throws<InvalidOperationException>(action);

            Assert.Contains("Base URL", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, handler.CallCount);
        }

        private sealed class RecordingHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> responder;

            public RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
            {
                this.responder = responder;
            }

            public HttpRequestMessage LastRequest { get; private set; }

            public string LastBody { get; private set; } = string.Empty;

            public int CallCount { get; private set; }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                LastRequest = request;
                LastBody = request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                return Task.FromResult(responder(request));
            }
        }
    }
}
