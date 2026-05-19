using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.ExcelAddIn.Updates;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class UpdateManifestClientTests
    {
        [Fact]
        public async Task GetManifestAsyncParsesJsonEvenWhenContentTypeIsOctetStream()
        {
            const string body = "{\"latestVersion\":\"1.0.176\",\"downloadUrl\":\"https://updates.example/download.exe\",\"releaseNotesUrl\":\"https://updates.example/notes\",\"publishedAtUtc\":\"2026-05-19T08:00:00Z\",\"title\":\"Release\",\"summary\":\"Summary\"}";
            using (var content = new ByteArrayContent(Encoding.UTF8.GetBytes(body)))
            {
                content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                var handler = new StubHttpHandler(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = content,
                });
                using (var httpClient = new HttpClient(handler))
                {
                    var client = new UpdateManifestClient(httpClient);

                    var manifest = await client.GetManifestAsync("https://updates.example/manifest.json", CancellationToken.None);

                    Assert.Equal("1.0.176", manifest.LatestVersion);
                    Assert.Equal("https://updates.example/download.exe", manifest.DownloadUrl);
                    Assert.Equal("https://updates.example/notes", manifest.ReleaseNotesUrl);
                    Assert.Equal(new DateTime(2026, 5, 19, 8, 0, 0, DateTimeKind.Utc), manifest.PublishedAtUtc);
                    Assert.Equal("Release", manifest.Title);
                    Assert.Equal("Summary", manifest.Summary);
                    Assert.Equal(1, handler.CallCount);
                }
            }
        }

        [Theory]
        [InlineData("{\"downloadUrl\":\"https://updates.example/download.exe\"}")]
        [InlineData("{\"latestVersion\":\"\"}")]
        [InlineData("not-json")]
        public async Task GetManifestAsyncRejectsInvalidManifests(string body)
        {
            var handler = new StubHttpHandler(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json"),
            });
            using (var httpClient = new HttpClient(handler))
            {
                var client = new UpdateManifestClient(httpClient);

                await Assert.ThrowsAsync<InvalidOperationException>(
                    () => client.GetManifestAsync("https://updates.example/manifest.json", CancellationToken.None));
            }
        }

        [Fact]
        public async Task GetManifestAsyncRejectsNonSuccessResponses()
        {
            var handler = new StubHttpHandler(new HttpResponseMessage(HttpStatusCode.InternalServerError));
            using (var httpClient = new HttpClient(handler))
            {
                var client = new UpdateManifestClient(httpClient);

                await Assert.ThrowsAsync<HttpRequestException>(
                    () => client.GetManifestAsync("https://updates.example/manifest.json", CancellationToken.None));
            }
        }

        private sealed class StubHttpHandler : HttpMessageHandler
        {
            private readonly HttpResponseMessage response;

            public StubHttpHandler(HttpResponseMessage response)
            {
                this.response = response;
            }

            public int CallCount { get; private set; }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                return Task.FromResult(response);
            }
        }
    }
}
