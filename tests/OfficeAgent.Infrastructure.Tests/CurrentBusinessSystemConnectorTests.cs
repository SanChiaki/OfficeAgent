using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class CurrentBusinessSystemConnectorTests
    {
        [Fact]
        public void CreateBindingSeedUsesConfigurableDefaults()
        {
            var connector = new CurrentBusinessSystemConnector(
                () => new AppSettings { BusinessBaseUrl = "https://business.internal.example" },
                new HttpClient(new RecordingHandler()));

            var project = new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "performance",
                DisplayName = "绩效项目",
            };
            var binding = connector.CreateBindingSeed("Sheet1", project);

            Assert.Equal("Sheet1", binding.SheetName);
            Assert.Equal("current-business-system", binding.SystemKey);
            Assert.Equal("performance", binding.ProjectId);
            Assert.Equal("绩效项目", binding.ProjectName);
            Assert.Equal(1, binding.HeaderStartRow);
            Assert.Equal(2, binding.HeaderRowCount);
            Assert.Equal(3, binding.DataStartRow);
        }

        [Fact]
        public void GetProjectsCallsProjectsEndpoint()
        {
            var handler = new RecordingHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            var projects = connector.GetProjects();

            Assert.Equal("/projects", handler.LastPath);
            Assert.Equal("https://api.internal.example/projects", handler.LastUri);
            Assert.Single(projects);
            Assert.Equal("current-business-system", projects[0].SystemKey);
            Assert.Equal("performance", projects[0].ProjectId);
        }

        [Theory]
        [InlineData(HttpStatusCode.Unauthorized)]
        [InlineData(HttpStatusCode.Forbidden)]
        public void GetProjectsPromptsLoginWhenProjectsEndpointRequiresAuthentication(HttpStatusCode statusCode)
        {
            var connector = CurrentBusinessSystemConnector.ForTests(
                "https://api.internal.example",
                new AuthRequiredProjectsHandler(statusCode));

            var error = Assert.Throws<AuthenticationRequiredException>(() => connector.GetProjects());

            Assert.Equal("当前未登录，请先登录", error.Message);
        }

        [Fact]
        public void GetFieldMappingDefinitionExposesCurrentSystemSemanticRoles()
        {
            var connector = new CurrentBusinessSystemConnector(
                () => new AppSettings { BusinessBaseUrl = "https://business.internal.example" },
                new HttpClient(new RecordingHandler()));

            var definition = connector.GetFieldMappingDefinition("performance");
            Assert.Equal("current-business-system", definition.SystemKey);
            Assert.Equal(
                new[]
                {
                    (CurrentBusinessFieldMappingHeaders.HeaderType, FieldMappingSemanticRole.HeaderType, string.Empty),
                    (CurrentBusinessFieldMappingHeaders.IsdpLevel1, FieldMappingSemanticRole.DefaultSingleHeaderText, CurrentBusinessFieldMappingColumns.DefaultLevel1),
                    (CurrentBusinessFieldMappingHeaders.IsdpLevel1, FieldMappingSemanticRole.DefaultParentHeaderText, CurrentBusinessFieldMappingColumns.DefaultLevel1),
                    (CurrentBusinessFieldMappingHeaders.IsdpLevel2, FieldMappingSemanticRole.DefaultChildHeaderText, CurrentBusinessFieldMappingColumns.DefaultLevel2),
                    (CurrentBusinessFieldMappingHeaders.ExcelLevel1, FieldMappingSemanticRole.CurrentSingleHeaderText, CurrentBusinessFieldMappingColumns.CurrentLevel1),
                    (CurrentBusinessFieldMappingHeaders.ExcelLevel1, FieldMappingSemanticRole.CurrentParentHeaderText, CurrentBusinessFieldMappingColumns.CurrentLevel1),
                    (CurrentBusinessFieldMappingHeaders.ExcelLevel2, FieldMappingSemanticRole.CurrentChildHeaderText, CurrentBusinessFieldMappingColumns.CurrentLevel2),
                    (CurrentBusinessFieldMappingHeaders.HeaderId, FieldMappingSemanticRole.HeaderIdentity, string.Empty),
                    (CurrentBusinessFieldMappingHeaders.ApiFieldKey, FieldMappingSemanticRole.ApiFieldKey, string.Empty),
                    (CurrentBusinessFieldMappingHeaders.IsIdColumn, FieldMappingSemanticRole.IsIdColumn, string.Empty),
                    (CurrentBusinessFieldMappingHeaders.ActivityId, FieldMappingSemanticRole.ActivityIdentity, string.Empty),
                    (CurrentBusinessFieldMappingHeaders.PropertyId, FieldMappingSemanticRole.PropertyIdentity, string.Empty),
                },
                definition.Columns.Select(column => (column.ColumnName, column.Role, column.RoleKey)).ToArray());
        }

        [Fact]
        public void GetFieldMappingDefinitionAllowsDynamicProjectId()
        {
            var connector = new CurrentBusinessSystemConnector(
                () => new AppSettings { BusinessBaseUrl = "https://business.internal.example" },
                new HttpClient(new RecordingHandler()));

            var definition = connector.GetFieldMappingDefinition("other-project");

            Assert.Equal("current-business-system", definition.SystemKey);
        }

        [Theory]
        [InlineData("")]
        [InlineData("  ")]
        public void GetFieldMappingDefinitionRejectsBlankProjectId(string projectId)
        {
            var connector = new CurrentBusinessSystemConnector(
                () => new AppSettings { BusinessBaseUrl = "https://business.internal.example" },
                new HttpClient(new RecordingHandler()));

            var error = Assert.Throws<InvalidOperationException>(() => connector.GetFieldMappingDefinition(projectId));

            Assert.Contains("Project id is required", error.Message);
        }

        [Theory]
        [InlineData("")]
        [InlineData("  ")]
        public void BuildFieldMappingSeedRejectsBlankProjectId(string projectId)
        {
            var handler = new SequencedResponseHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            var error = Assert.Throws<InvalidOperationException>(() => connector.BuildFieldMappingSeed("Sheet1", projectId));

            Assert.Contains("Project id is required", error.Message);
            Assert.Empty(handler.Requests);
        }

        [Fact]
        public void BuildFieldMappingSeedCallsHeadAndFindBeforeReturningRows()
        {
            var handler = new SequencedResponseHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            var rows = connector.BuildFieldMappingSeed("Sheet1", "performance");

            Assert.Equal(2, handler.Requests.Count);
            Assert.Equal("/head", handler.Requests[0].Path);
            Assert.Contains("\"projectId\":\"performance\"", handler.Requests[0].Body);
            Assert.Equal("/find", handler.Requests[1].Path);
            Assert.Contains("\"projectId\":\"performance\"", handler.Requests[1].Body);
            Assert.Contains("\"ids\":[]", handler.Requests[1].Body);
            Assert.Contains("\"rowIds\":[]", handler.Requests[1].Body);
            Assert.Contains("\"fieldKeys\":[]", handler.Requests[1].Body);
            Assert.NotEmpty(rows);
        }

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
        public void GetProjectsRequiresBusinessBaseUrlWhenInvoked()
        {
            var connector = new CurrentBusinessSystemConnector(
                () => new AppSettings
                {
                    BaseUrl = "https://llm.internal.example",
                    BusinessBaseUrl = string.Empty,
                });

            var error = Assert.Throws<InvalidOperationException>(() => connector.GetProjects());

            Assert.Contains("Business API Base URL", error.Message);
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
        public void FindSendsIdsAndLegacyRowIdsForCompatibility()
        {
            var handler = new RecordingHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            connector.Find("performance", new[] { "row-1" }, new[] { "name" });

            Assert.Contains("\"ids\":[\"row-1\"]", handler.LastBody);
            Assert.Contains("\"rowIds\":[\"row-1\"]", handler.LastBody);
        }

        [Fact]
        public void BatchSaveRetriesWithLegacyItemsWrapperWhenArrayPayloadIsRejected()
        {
            var handler = new LegacyBatchSaveHandler();
            var connector = CurrentBusinessSystemConnector.ForTests("https://api.internal.example", handler);

            connector.BatchSave(
                "performance",
                new[]
                {
                    new CellChange { RowId = "row-1", ApiFieldKey = "name", NewValue = "A" },
                });

            Assert.Equal(2, handler.CallCount);
            Assert.StartsWith("[", handler.RequestBodies[0].TrimStart(), StringComparison.Ordinal);
            Assert.Contains("\"items\":[", handler.RequestBodies[1]);
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

                var responseBody = request.RequestUri?.AbsolutePath switch
                {
                    "/projects" => @"[{""projectId"":""performance"",""displayName"":""绩效项目""}]",
                    _ => "[]",
                };

                var response = new HttpResponseMessage(System.Net.HttpStatusCode.OK)
                {
                    Content = new StringContent(responseBody, Encoding.UTF8, "application/json"),
                };

                return Task.FromResult(response);
            }
        }

        private sealed class AuthRequiredProjectsHandler : HttpMessageHandler
        {
            private readonly HttpStatusCode statusCode;

            public AuthRequiredProjectsHandler(HttpStatusCode statusCode)
            {
                this.statusCode = statusCode;
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                return Task.FromResult(new HttpResponseMessage(statusCode)
                {
                    Content = new StringContent(
                        "{\"code\":\"unauthorized\",\"message\":\"未登录，请先通过 SSO 登录。\"}",
                        Encoding.UTF8,
                        "application/json"),
                });
            }
        }

        private sealed class LegacyBatchSaveHandler : HttpMessageHandler
        {
            public List<string> RequestBodies { get; } = new List<string>();
            public int CallCount => RequestBodies.Count;

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var body = request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                RequestBodies.Add(body);

                if (RequestBodies.Count == 1)
                {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
                    {
                        Content = new StringContent("{\"code\":\"bad_request\",\"message\":\"items 必须为非空数组。\"}", Encoding.UTF8, "application/json"),
                    });
                }

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("{}", Encoding.UTF8, "application/json"),
                });
            }
        }

        private sealed class SequencedResponseHandler : HttpMessageHandler
        {
            public List<(string Path, string Body)> Requests { get; } = new List<(string Path, string Body)>();

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var body = request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                Requests.Add((request.RequestUri?.AbsolutePath ?? string.Empty, body));

                var responseBody = request.RequestUri?.AbsolutePath switch
                {
                    "/head" => @"{""headList"":[{""fieldKey"":""row_id"",""headerText"":""ID"",""headType"":""single"",""isId"":true},{""headType"":""activity"",""activityId"":""12345678"",""activityName"":""测试活动111""}]}",
                    "/find" => @"[{""start_12345678"":""2026-01-02"",""end_12345678"":""2026-01-03""}]",
                    _ => "[]",
                };

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(responseBody, Encoding.UTF8, "application/json"),
                });
            }
        }
    }
}
