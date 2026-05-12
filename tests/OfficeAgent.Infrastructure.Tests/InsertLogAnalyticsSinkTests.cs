using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Analytics;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Analytics;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class InsertLogAnalyticsSinkTests
    {
        [Fact]
        public async Task WriteAsyncPostsInsertLogEnvelopeWithJsonAnswer()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK));
            var sink = new InsertLogAnalyticsSink(
                () => new AppSettings { AnalyticsBaseUrl = "https://analytics.internal.example/v1/" },
                new HttpClient(handler));
            var analyticsEvent = new AnalyticsEvent
            {
                EventName = "ribbon.download.clicked",
                Source = "ribbon",
                Properties = new Dictionary<string, object>(StringComparer.Ordinal)
                {
                    ["projectId"] = "performance",
                    ["projectName"] = "\u7EE9\u6548\u9879\u76EE",
                },
            };

            await sink.WriteAsync(analyticsEvent, CancellationToken.None);

            Assert.Equal("https://analytics.internal.example/v1/insertLog", handler.LastRequest.RequestUri.ToString());
            var envelope = JObject.Parse(handler.LastRequestBody);
            Assert.Equal("excelAi", (string)envelope["frontEndIntent"]);
            Assert.Equal("Excel", (string)envelope["clientSource"]);
            Assert.Equal(1, (int)envelope["questionType"]);
            Assert.False(string.IsNullOrWhiteSpace((string)envelope["askId"]));
            Assert.False(string.IsNullOrWhiteSpace((string)envelope["talkId"]));

            var answer = Assert.IsType<string>(envelope["answer"]?.ToObject<object>());
            var answerJson = JObject.Parse(answer);
            Assert.Equal("ribbon.download.clicked", (string)answerJson["eventName"]);
            Assert.Equal("\u7EE9\u6548\u9879\u76EE", (string)answerJson["properties"]?["projectName"]);
        }

        [Fact]
        public async Task WriteAsyncRejectsMissingAnalyticsBaseUrl()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK));
            var sink = new InsertLogAnalyticsSink(
                () => new AppSettings { AnalyticsBaseUrl = " " },
                new HttpClient(handler));

            var error = await Assert.ThrowsAsync<InvalidOperationException>(
                () => sink.WriteAsync(new AnalyticsEvent(), CancellationToken.None));

            Assert.Equal("The configured Analytics Base URL is invalid. Update settings and try again.", error.Message);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public async Task WriteAsyncThrowsForNonSuccessResponse()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent("bad request"),
            });
            var sink = new InsertLogAnalyticsSink(
                () => new AppSettings { AnalyticsBaseUrl = "https://analytics.internal.example" },
                new HttpClient(handler));

            var error = await Assert.ThrowsAsync<InvalidOperationException>(
                () => sink.WriteAsync(new AnalyticsEvent(), CancellationToken.None));

            Assert.Contains("Analytics request failed (400 Bad Request): bad request", error.Message);
        }

        private sealed class RecordingHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> responder;

            public RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
            {
                this.responder = responder;
            }

            public HttpRequestMessage LastRequest { get; private set; }

            public string LastRequestBody { get; private set; }

            public int CallCount { get; private set; }

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                LastRequest = request;
                LastRequestBody = await (request.Content?.ReadAsStringAsync() ?? Task.FromResult(string.Empty));
                return responder(request);
            }
        }
    }
}
