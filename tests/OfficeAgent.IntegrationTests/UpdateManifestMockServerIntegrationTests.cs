using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Xunit;

namespace OfficeAgent.IntegrationTests
{
    [Collection(MockServerCollection.Name)]
    public sealed class UpdateManifestMockServerIntegrationTests
    {
        private readonly MockServerFixture fixture;

        public UpdateManifestMockServerIntegrationTests(MockServerFixture fixture)
        {
            this.fixture = fixture;
        }

        [Fact]
        public async Task UpdateManifestEndpointReturnsOctetStreamJsonWithoutAuthentication()
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };

            using var response = await client.GetAsync(fixture.BusinessUrl + "/update-manifest");

            response.EnsureSuccessStatusCode();
            Assert.Equal("application/octet-stream", response.Content.Headers.ContentType?.MediaType);
            var body = await response.Content.ReadAsStringAsync();
            var manifest = JObject.Parse(body);
            Assert.Equal("9.9.9", manifest.Value<string>("latestVersion"));
            Assert.Equal("http://localhost:3200/update-download/xISDP.Setup.exe", manifest.Value<string>("downloadUrl"));
            Assert.Equal("http://localhost:3200/update-release-notes", manifest.Value<string>("releaseNotesUrl"));
            Assert.False(string.IsNullOrWhiteSpace(manifest.Value<string>("title")));
            Assert.False(string.IsNullOrWhiteSpace(manifest.Value<string>("summary")));
        }
    }
}
