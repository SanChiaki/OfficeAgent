using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class CurrentBusinessSystemConnectorTests
    {
        [Fact]
        public void FindUsesBusinessBaseUrlInsteadOfTheLlmBaseUrl()
        {
            var handler = new RecordingHandler();
            var connector = new CurrentBusinessSystemConnector(
                () => new AppSettings
                {
                    BaseUrl = "https://llm.internal.example",
                    BusinessBaseUrl = "https://business.internal.example",
                },
                new HttpClient(handler));

            connector.Find("performance", Array.Empty<string>(), Array.Empty<string>());

            Assert.Equal("/find", handler.LastPath);
            Assert.Equal("https://business.internal.example/find", handler.LastUri);
        }

        [Fact]
        public void GetProjectsDoesNotRequireBusinessBaseUrlDuringConstruction()
        {
            var connector = new CurrentBusinessSystemConnector(
                () => new AppSettings
                {
                    BaseUrl = "https://llm.internal.example",
                    BusinessBaseUrl = string.Empty,
                });

            var projects = connector.GetProjects();

            Assert.Single(projects);
            Assert.Equal("performance", projects[0].ProjectId);
        }

        [Fact]
        public void BatchSaveSendsOneItemPerChangedCell()
        {
            var handler = new RecordingHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            connector.BatchSave(
                "performance",
                new[]
                {
                    new CellChange { RowId = "row-1", ApiFieldKey = "name", NewValue = "A" },
                    new CellChange { RowId = "row-1", ApiFieldKey = "start_12345678", NewValue = "2026-01-02" },
                });

            Assert.Equal("/batchSave", handler.LastPath);
            Assert.Contains("\"ProjectId\":\"performance\"", handler.LastBody);
            Assert.Equal(2, handler.ItemCount);
        }

        [Fact]
        public void BatchSaveShortCircuitsWhenChangesEmpty()
        {
            var handler = new RecordingHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            connector.BatchSave("performance", Array.Empty<CellChange>());

            Assert.Equal(0, handler.CallCount);
            Assert.Equal(string.Empty, handler.LastPath);
            Assert.Equal(string.Empty, handler.LastBody);
        }

        private sealed class RecordingHandler : HttpMessageHandler
        {
            public string LastPath { get; private set; } = string.Empty;
            public string LastUri { get; private set; } = string.Empty;
            public string LastBody { get; private set; } = string.Empty;
            public int ItemCount { get; private set; }
            public int CallCount { get; private set; }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                LastPath = request.RequestUri?.AbsolutePath ?? string.Empty;
                LastUri = request.RequestUri?.ToString() ?? string.Empty;
                LastBody = request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                ItemCount = 0;
                if (!string.IsNullOrEmpty(LastBody) && LastBody.TrimStart().StartsWith("[", StringComparison.Ordinal))
                {
                    var items = JArray.Parse(LastBody);
                    ItemCount = items.Count;
                }

                var response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
                {
                    Content = new StringContent("[]", Encoding.UTF8, "application/json"),
                };

                return Task.FromResult(response);
            }
        }
    }
}
