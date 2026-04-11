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
    public sealed class BusinessApiClientTests
    {
        private const string ProjectA = "\u9879\u76EEA";

        [Fact]
        public void UploadUsesBusinessBaseUrlInsteadOfTheLlmBaseUrl()
        {
            var handler = new RecordingHandler(_ => CreateSuccessResponse());
            var client = new BusinessApiClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://llm.internal.example",
                    BusinessBaseUrl = "https://business.internal.example",
                    Model = "gpt-5-mini",
                });

            client.Upload(CreatePreview());

            Assert.Equal("https://business.internal.example/upload_data", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void UploadFallsBackToTheDefaultBaseUrlWhenSettingsLeaveItBlank()
        {
            var handler = new RecordingHandler(_ => CreateSuccessResponse());
            var client = new BusinessApiClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "   ",
                    BusinessBaseUrl = "   ",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Upload(CreatePreview()));

            Assert.Equal("The configured Business API Base URL is invalid. Update settings and try again.", error.Message);
        }

        [Fact]
        public void UploadNormalizesTheConfiguredBaseUrlAndSendsTheApiKey()
        {
            var handler = new RecordingHandler(_ => CreateSuccessResponse());
            var client = new BusinessApiClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BusinessBaseUrl = " https://api.internal.example/ ",
                    Model = "gpt-5-mini",
                });

            var result = client.Upload(CreatePreview());

            Assert.Equal("https://api.internal.example/upload_data", handler.LastRequest.RequestUri.ToString());
            Assert.Equal("Bearer", handler.LastRequest.Headers.Authorization?.Scheme);
            Assert.Equal("secret-token", handler.LastRequest.Headers.Authorization?.Parameter);
            Assert.Equal("Uploaded 2 row(s).", result.Message);
        }

        [Fact]
        public void UploadFormatsStructuredApiErrorsFromJsonResponseBodies()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = new StringContent("{\"code\":\"invalid_project\",\"message\":\"Project not found\"}"),
            });
            var client = new BusinessApiClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BusinessBaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Upload(CreatePreview()));

            Assert.Equal("Business API request failed (400 invalid_project): Project not found", error.Message);
        }

        [Fact]
        public void UploadRetriesOnceForRetryableResponsesBeforeSucceeding()
        {
            var callCount = 0;
            var handler = new RecordingHandler(_ =>
            {
                callCount++;
                if (callCount == 1)
                {
                    return new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                    {
                        Content = new StringContent("temporarily unavailable"),
                    };
                }

                return CreateSuccessResponse();
            });
            var client = new BusinessApiClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BusinessBaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var result = client.Upload(CreatePreview());

            Assert.Equal(2, handler.CallCount);
            Assert.Equal("Uploaded 2 row(s).", result.Message);
        }

        [Fact]
        public void UploadFormatsTimeoutFailuresAsAControlledErrorWithoutRetrying()
        {
            var handler = new RecordingHandler(_ => throw new TaskCanceledException("The request timed out."));
            var httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(15),
            };
            var client = new BusinessApiClient(
                httpClient,
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BusinessBaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Upload(CreatePreview()));

            Assert.Equal(1, handler.CallCount);
            Assert.Equal("Business API request timed out after 15 seconds.", error.Message);
        }

        [Fact]
        public void UploadRejectsInvalidBaseUrlsWithAControlledError()
        {
            var handler = new RecordingHandler(_ => CreateSuccessResponse());
            var client = new BusinessApiClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BusinessBaseUrl = "api.internal.example",
                    Model = "gpt-5-mini",
                });

            var error = Assert.Throws<InvalidOperationException>(() => client.Upload(CreatePreview()));

            Assert.Equal("The configured Business API Base URL is invalid. Update settings and try again.", error.Message);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public void UploadPreservesBaseUrlPathPrefixesWhenBuildingTheEndpoint()
        {
            var handler = new RecordingHandler(_ => CreateSuccessResponse());
            var client = new BusinessApiClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BusinessBaseUrl = "https://api.internal.example/v1/",
                    Model = "gpt-5-mini",
                });

            client.Upload(CreatePreview());

            Assert.Equal("https://api.internal.example/v1/upload_data", handler.LastRequest.RequestUri.ToString());
        }

        private static UploadPreview CreatePreview()
        {
            return new UploadPreview
            {
                ProjectName = ProjectA,
                SheetName = "Sheet1",
                Address = "A1:C3",
                Headers = new[] { "Name", "Region" },
                Rows = new[]
                {
                    new[] { "Project A", "CN" },
                    new[] { "Project B", "US" },
                },
                Records = new[]
                {
                    new System.Collections.Generic.Dictionary<string, string>
                    {
                        ["Name"] = "Project A",
                        ["Region"] = "CN",
                    },
                    new System.Collections.Generic.Dictionary<string, string>
                    {
                        ["Name"] = "Project B",
                        ["Region"] = "US",
                    },
                },
            };
        }

        private static HttpResponseMessage CreateSuccessResponse()
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"savedCount\":2,\"message\":\"Uploaded 2 row(s).\"}"),
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

            public int CallCount { get; private set; }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                LastRequest = request;
                return Task.FromResult(responder(request));
            }
        }
    }
}
