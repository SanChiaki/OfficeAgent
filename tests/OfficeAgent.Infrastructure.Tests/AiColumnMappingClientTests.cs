using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class AiColumnMappingClientTests
    {
        [Fact]
        public void MapPostsColumnMappingRequestsToConfiguredChatCompletionsEndpoint()
        {
            var handler = new RecordingHandler(_ => CreateChatCompletionResponse(
                "{\"Mappings\":[{\"ExcelColumn\":2,\"ActualL1\":\"项目负责人\",\"TargetHeaderId\":\"owner_name\",\"TargetApiFieldKey\":\"owner_name\",\"Confidence\":0.91,\"Reason\":\"same meaning\"}],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = " https://api.internal.example/ ",
                    Model = "gpt-5-mini",
                });

            var result = client.Map(CreateRequest());

            Assert.Equal("https://api.internal.example/v1/chat/completions", handler.LastRequest.RequestUri.ToString());
            Assert.Equal("Bearer", handler.LastRequest.Headers.Authorization?.Scheme);
            Assert.Equal("secret-token", handler.LastRequest.Headers.Authorization?.Parameter);
            Assert.Contains("gpt-5-mini", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"response_format\":{\"type\":\"json_object\"}", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("项目负责人", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("owner_name", handler.LastBody, StringComparison.Ordinal);
            Assert.Equal(0.91, Assert.Single(result.Mappings).Confidence);
            Assert.Equal("same meaning", result.Mappings[0].Reason);
        }

        [Fact]
        public void MapPreservesBaseUrlPathPrefixes()
        {
            var handler = new RecordingHandler(_ => CreateChatCompletionResponse("{\"Mappings\":[],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example/v1/",
                    Model = "gpt-5-mini",
                });

            client.Map(CreateRequest());

            Assert.Equal("https://api.internal.example/v1/chat/completions", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void MapPostsAnthropicMessagesRequestsWhenConfigured()
        {
            var handler = new RecordingHandler(_ => CreateAnthropicMessageResponse(
                "{\"Mappings\":[{\"ExcelColumn\":2,\"ActualL1\":\"项目负责人\",\"TargetHeaderId\":\"owner_name\",\"TargetApiFieldKey\":\"owner_name\",\"Confidence\":0.91,\"Reason\":\"same meaning\"}],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = " https://api.anthropic.com/ ",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            var result = client.Map(CreateRequest());

            Assert.Equal("https://api.anthropic.com/v1/messages", handler.LastRequest.RequestUri.ToString());
            Assert.Null(handler.LastRequest.Headers.Authorization);
            Assert.True(handler.LastRequest.Headers.TryGetValues("x-api-key", out var apiKeyValues));
            Assert.Contains("secret-token", apiKeyValues);
            Assert.True(handler.LastRequest.Headers.TryGetValues("anthropic-version", out var versionValues));
            Assert.Contains("2023-06-01", versionValues);
            Assert.Contains("\"system\":", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"max_tokens\":4096", handler.LastBody, StringComparison.Ordinal);
            Assert.DoesNotContain("response_format", handler.LastBody, StringComparison.Ordinal);
            Assert.Equal("owner_name", Assert.Single(result.Mappings).TargetHeaderId);
        }

        [Fact]
        public void MapReadsStreamingAnthropicMessagesDeltas()
        {
            var handler = new RecordingHandler(_ => CreateStreamingChatCompletionResponse(
                "event: message_start\n"
                + "data: {\"type\":\"message_start\",\"message\":{\"id\":\"msg_123\"}}\n\n"
                + "event: content_block_delta\n"
                + "data: {\"type\":\"content_block_delta\",\"index\":0,\"delta\":{\"type\":\"text_delta\",\"text\":\"{\\\"Mappings\\\":[]\"}}\n\n"
                + "event: content_block_delta\n"
                + "data: {\"type\":\"content_block_delta\",\"index\":0,\"delta\":{\"type\":\"text_delta\",\"text\":\",\\\"Unmatched\\\":[]}\"}}\n\n"
                + "event: message_stop\n"
                + "data: {\"type\":\"message_stop\"}\n\n"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.anthropic.com",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            var result = client.Map(CreateRequest());

            Assert.Equal("https://api.anthropic.com/v1/messages", handler.LastRequest.RequestUri.ToString());
            Assert.Contains("\"stream\":true", handler.LastBody, StringComparison.Ordinal);
            Assert.Empty(result.Mappings);
            Assert.Empty(result.Unmatched);
        }

        [Fact]
        public void MapAppendsAnthropicVersionPathAfterProxyPrefixes()
        {
            var handler = new RecordingHandler(_ => CreateAnthropicMessageResponse("{\"Mappings\":[],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example/anthropic/",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            client.Map(CreateRequest());

            Assert.Equal("https://api.internal.example/anthropic/v1/messages", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void MapDoesNotDuplicateAnthropicVersionPath()
        {
            var handler = new RecordingHandler(_ => CreateAnthropicMessageResponse("{\"Mappings\":[],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example/anthropic/v1/",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            client.Map(CreateRequest());

            Assert.Equal("https://api.internal.example/anthropic/v1/messages", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void MapDoesNotDuplicateRootAnthropicVersionPath()
        {
            var handler = new RecordingHandler(_ => CreateAnthropicMessageResponse("{\"Mappings\":[],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.anthropic.com/v1/",
                    Model = "claude-3-5-sonnet-latest",
                    ApiFormat = "anthropic-messages",
                });

            client.Map(CreateRequest());

            Assert.Equal("https://api.anthropic.com/v1/messages", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void MapUsesCompactJsonForColumnMappingPrompt()
        {
            var handler = new RecordingHandler(_ => CreateChatCompletionResponse("{\"Mappings\":[],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            client.Map(CreateRequest());

            Assert.Contains("Column mapping request:\\n{\\\"SystemKey\\\":", handler.LastBody, StringComparison.Ordinal);
            Assert.DoesNotContain("\\n  \\\"SystemKey\\\"", handler.LastBody, StringComparison.Ordinal);
        }

        [Fact]
        public void MapReadsStreamingChatCompletionDeltas()
        {
            var handler = new RecordingHandler(_ => CreateStreamingChatCompletionResponse(
                "data: {\"choices\":[{\"delta\":{\"content\":\"{\\\"Mappings\\\":[\"}}]}\n\n"
                + "data: {\"choices\":[{\"delta\":{\"content\":\"{\\\"ExcelColumn\\\":2,\\\"ActualL1\\\":\\\"项目负责人\\\",\\\"TargetHeaderId\\\":\\\"owner_name\\\",\\\"TargetApiFieldKey\\\":\\\"owner_name\\\",\\\"Confidence\\\":0.91}\"}}]}\n\n"
                + "data: {\"choices\":[{\"delta\":{\"content\":\"],\\\"Unmatched\\\":[]}\"}}]}\n\n"
                + "data: [DONE]\n\n"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var result = client.Map(CreateRequest());

            Assert.Contains("\"stream\":true", handler.LastBody, StringComparison.Ordinal);
            var mapping = Assert.Single(result.Mappings);
            Assert.Equal(2, mapping.ExcelColumn);
            Assert.Equal("owner_name", mapping.TargetHeaderId);
            Assert.Equal(0.91, mapping.Confidence);
        }

        [Fact]
        public void MapIgnoresStreamingChunksWithoutChoices()
        {
            var handler = new RecordingHandler(_ => CreateStreamingChatCompletionResponse(
                "data: {\"choices\":[{\"delta\":{\"content\":\"{\\\"Mappings\\\":[],\\\"Unmatched\\\":\"}}]}\n\n"
                + "data: {\"choices\":[],\"usage\":{\"prompt_tokens\":10,\"completion_tokens\":5}}\n\n"
                + "data: {\"choices\":[{\"delta\":{\"content\":\"[]}\"},\"finish_reason\":\"stop\"}]}\n\n"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var result = client.Map(CreateRequest());

            Assert.Empty(result.Mappings);
            Assert.Empty(result.Unmatched);
        }

        [Fact]
        public void MapCompletesStreamingResponseWhenFinishReasonIsStopWithoutDoneMarker()
        {
            var handler = new RecordingHandler(_ => CreateStreamingChatCompletionResponse(
                "data: {\"choices\":[{\"delta\":{\"content\":\"{\\\"Mappings\\\":[],\\\"Unmatched\\\":[]}\"},\"finish_reason\":\"stop\"}]}\n\n"
                + "data: {\"choices\":[{\"delta\":{\"content\":\"this should not be appended\"}}]}\n\n"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var result = client.Map(CreateRequest());

            Assert.Empty(result.Mappings);
            Assert.Empty(result.Unmatched);
        }

        [Fact]
        public void MapFallsBackToNonStreamingWhenStreamingRequestIsRejected()
        {
            RecordingHandler handler = null;
            handler = new RecordingHandler(request =>
            {
                if (handler.CallCount == 1)
                {
                    return new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent("{\"error\":{\"message\":\"stream unsupported\"}}"),
                    };
                }

                return CreateChatCompletionResponse(
                    "{\"Mappings\":[{\"ExcelColumn\":2,\"ActualL1\":\"项目负责人\",\"TargetHeaderId\":\"owner_name\",\"TargetApiFieldKey\":\"owner_name\",\"Confidence\":0.91}],\"Unmatched\":[]}");
            });
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var result = client.Map(CreateRequest());

            Assert.Equal(2, handler.CallCount);
            Assert.Contains("\"stream\":false", handler.LastBody, StringComparison.Ordinal);
            Assert.Equal("owner_name", Assert.Single(result.Mappings).TargetHeaderId);
        }

        [Fact]
        public void MapParsesTextFromContentArrays()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"choices\":[{"
                    + "\"message\":{"
                    + "\"content\":["
                    + "{\"type\":\"output_text\",\"text\":\"{\\\"Mappings\\\":[{\\\"ExcelColumn\\\":3,\\\"ActualL1\\\":\\\"状态\\\",\\\"TargetHeaderId\\\":\\\"status\\\",\\\"TargetApiFieldKey\\\":\\\"status\\\",\\\"Confidence\\\":0.88}],\\\"Unmatched\\\":[]}\"}"
                    + "]"
                    + "}"
                    + "}]"
                    + "}"),
            });
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var result = client.Map(CreateRequest());

            var mapping = Assert.Single(result.Mappings);
            Assert.Equal(3, mapping.ExcelColumn);
            Assert.Equal("status", mapping.TargetApiFieldKey);
        }

        [Fact]
        public async Task MapAsyncUsesTheSameResponseParsingAsMap()
        {
            var handler = new RecordingHandler(_ => CreateChatCompletionResponse(
                "{\"Mappings\":[],\"Unmatched\":[{\"ExcelColumn\":4,\"ActualL1\":\"备注\",\"Reason\":\"no candidate\"}]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var result = await client.MapAsync(CreateRequest());

            Assert.Empty(result.Mappings);
            Assert.Equal("no candidate", Assert.Single(result.Unmatched).Reason);
        }

        [Fact]
        public void MapRejectsMissingApiKeys()
        {
            var handler = new RecordingHandler(_ => CreateChatCompletionResponse("{\"Mappings\":[],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = " ",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Map(CreateRequest()));

            Assert.Contains("API Key", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public void MapRejectsInvalidBaseUrls()
        {
            var handler = new RecordingHandler(_ => CreateChatCompletionResponse("{\"Mappings\":[],\"Unmatched\":[]}"));
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "api.internal.example",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Map(CreateRequest()));

            Assert.Contains("Base URL", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public void MapFormatsHttpErrors()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent("{\"error\":{\"message\":\"bad request\"}}"),
            });
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Map(CreateRequest()));

            Assert.Contains("AI column mapping API request failed (400", error.Message, StringComparison.Ordinal);
            Assert.Contains("bad request", error.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("")]
        [InlineData("{\"choices\":[{\"message\":{\"content\":\"not json\"}}]}")]
        [InlineData("{\"choices\":[{\"message\":{\"content\":\"{\\\"Unmatched\\\":[]}\"}}]}")]
        [InlineData("{\"choices\":[{\"message\":{}}]}")]
        public void MapRejectsMalformedMappingResponses(string responseBody)
        {
            var handler = new RecordingHandler(_ => string.IsNullOrEmpty(responseBody)
                ? new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(string.Empty) }
                : new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(responseBody) });
            var client = new AiColumnMappingClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Map(CreateRequest()));

            Assert.Contains("AI column mapping API", error.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static AiColumnMappingRequest CreateRequest()
        {
            return new AiColumnMappingRequest
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ActualHeaders = new[]
                {
                    new AiColumnMappingActualHeader
                    {
                        ExcelColumn = 2,
                        ActualL1 = "项目负责人",
                        DisplayText = "项目负责人",
                    },
                },
                Candidates = new[]
                {
                    new AiColumnMappingCandidate
                    {
                        HeaderId = "owner_name",
                        ApiFieldKey = "owner_name",
                        HeaderType = "single",
                        IsdpL1 = "负责人",
                        CurrentExcelL1 = "负责人",
                    },
                },
            };
        }

        private static HttpResponseMessage CreateChatCompletionResponse(string contentJson)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(JObject.FromObject(new
                {
                    id = "chatcmpl-123",
                    @object = "chat.completion",
                    choices = new[]
                    {
                        new
                        {
                            index = 0,
                            message = new
                            {
                                role = "assistant",
                                content = contentJson,
                            },
                            finish_reason = "stop",
                        },
                    },
                }).ToString()),
            };
        }

        private static HttpResponseMessage CreateAnthropicMessageResponse(string contentJson)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(JObject.FromObject(new
                {
                    id = "msg_123",
                    type = "message",
                    role = "assistant",
                    content = new[]
                    {
                        new
                        {
                            type = "text",
                            text = contentJson,
                        },
                    },
                    stop_reason = "end_turn",
                }).ToString()),
            };
        }

        private static HttpResponseMessage CreateStreamingChatCompletionResponse(string content)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(content),
            };
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
