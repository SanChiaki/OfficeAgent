using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Authentication;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.IntegrationTests
{
    [Collection(MockServerCollection.Name)]
    public sealed class CurrentBusinessSystemConnectorIntegrationTests : IClassFixture<MockServerFixture>
    {
        private readonly MockServerFixture fixture;

        public CurrentBusinessSystemConnectorIntegrationTests(MockServerFixture fixture)
        {
            this.fixture = fixture;
        }

        [Fact]
        public async Task FindAndBatchSaveRoundTripAgainstMockServer()
        {
            var connector = await CreateConnectorAsync();
            var rows = connector.Find("performance", new[] { "row-1" }, new[] { "owner_name", "start_12345678" });
            Assert.Single(rows);

            connector.BatchSave("performance", new[]
            {
                new CellChange
                {
                    RowId = "row-1",
                    ApiFieldKey = "owner_name",
                    NewValue = "Updated",
                },
            });

            var updatedRows = connector.Find("performance", new[] { "row-1" }, new[] { "owner_name" });
            Assert.Single(updatedRows);
            Assert.Equal("Updated", updatedRows[0]["owner_name"]?.ToString());
        }

        [Fact]
        public async Task BuildFieldMappingSeedAndBatchSaveRoundTripAgainstMockServer()
        {
            var connector = await CreateConnectorAsync();

            var mappings = connector.BuildFieldMappingSeed("Sheet1", "performance");

            Assert.Contains(
                mappings,
                row => string.Equals(row.Values[CurrentBusinessFieldMappingColumns.ApiFieldKey], "start_12345678", StringComparison.Ordinal));
            Assert.Contains(
                mappings,
                row => string.Equals(row.Values[CurrentBusinessFieldMappingColumns.ApiFieldKey], "owner_name", StringComparison.Ordinal));

            var rows = connector.Find("performance", Array.Empty<string>(), new[] { "owner_name" });
            var row = Assert.Single(rows, item => string.Equals(item["row_id"]?.ToString(), "row-1", StringComparison.Ordinal));

            connector.BatchSave(
                "performance",
                new[]
                {
                    new CellChange
                    {
                        RowId = row["row_id"]?.ToString(),
                        ApiFieldKey = "owner_name",
                        NewValue = "李四",
                    },
                });

            var afterSave = connector.Find("performance", new[] { "row-1" }, new[] { "owner_name" });
            Assert.Equal("李四", afterSave.Single()["owner_name"]?.ToString());
        }

        [Fact]
        public async Task GetSchemaIncludesActivityColumns()
        {
            var connector = await CreateConnectorAsync();
            var schema = connector.GetSchema("performance");

            Assert.NotNull(schema);
            var columns = schema.Columns;
            Assert.Contains(columns, column => column.ApiFieldKey == "row_id" && column.IsIdColumn);
            Assert.Contains(columns, column => column.ApiFieldKey == "owner_name" && column.ColumnKind == WorksheetColumnKind.Single);
            Assert.Contains(columns, column => column.ApiFieldKey == "start_12345678" && column.ColumnKind == WorksheetColumnKind.ActivityProperty);
            Assert.Contains(columns, column => column.ApiFieldKey == "end_12345678" && column.ColumnKind == WorksheetColumnKind.ActivityProperty);
            var activityColumn = columns.First(column => column.ApiFieldKey == "start_12345678");
            Assert.Equal("测试活动111", activityColumn.ActivityName);
            Assert.Equal("开始时间", activityColumn.ChildHeaderText);
        }

        [Fact]
        public async Task GetProjectsReturnsAdditionalMockProjectsWithIndependentRows()
        {
            var connector = await CreateConnectorAsync();

            var projects = connector.GetProjects();

            Assert.Equal(4, projects.Count);
            Assert.Contains(
                projects,
                project => string.Equals(project.ProjectId, "performance", StringComparison.Ordinal)
                    && string.Equals(project.DisplayName, "绩效项目", StringComparison.Ordinal));
            Assert.Contains(
                projects,
                project => string.Equals(project.ProjectId, "delivery-tracker", StringComparison.Ordinal)
                    && string.Equals(project.DisplayName, "交付跟踪项目", StringComparison.Ordinal));
            Assert.Contains(
                projects,
                project => string.Equals(project.ProjectId, "customer-onboarding", StringComparison.Ordinal)
                    && string.Equals(project.DisplayName, "客户上线项目", StringComparison.Ordinal));
            Assert.Contains(
                projects,
                project => string.Equals(project.ProjectId, "large-activity-benchmark", StringComparison.Ordinal)
                    && string.Equals(project.DisplayName, "大数据活动压测项目", StringComparison.Ordinal));

            var rows = connector.Find("delivery-tracker", Array.Empty<string>(), Array.Empty<string>());

            Assert.Equal(10, rows.Count);
            Assert.All(
                rows,
                row => Assert.StartsWith(
                    "delivery-row-",
                    row["row_id"]?.ToString() ?? string.Empty,
                    StringComparison.Ordinal));
        }

        [Fact]
        public async Task GetProjectsIncludesLargeBenchmarkProjectWithActivitySchema()
        {
            var connector = await CreateConnectorAsync();

            var projects = connector.GetProjects();

            Assert.Equal(4, projects.Count);
            Assert.Contains(
                projects,
                project => string.Equals(project.ProjectId, "large-activity-benchmark", StringComparison.Ordinal)
                    && string.Equals(project.DisplayName, "大数据活动压测项目", StringComparison.Ordinal));

            var rows = connector.Find("large-activity-benchmark", Array.Empty<string>(), Array.Empty<string>());

            Assert.Equal(10000, rows.Count);

            var firstRow = rows[0];
            Assert.Equal(11, firstRow.Count);
            Assert.Equal("benchmark-row-00001", firstRow["row_id"]?.ToString());
            Assert.True(firstRow.ContainsKey("name_benchmarka"));
            Assert.True(firstRow.ContainsKey("start_benchmarka"));
            Assert.True(firstRow.ContainsKey("end_benchmarka"));
            Assert.True(firstRow.ContainsKey("name_benchmarkb"));
            Assert.True(firstRow.ContainsKey("start_benchmarkb"));
            Assert.True(firstRow.ContainsKey("end_benchmarkb"));

            var schema = connector.GetSchema("large-activity-benchmark");

            Assert.Contains(schema.Columns, column => column.ApiFieldKey == "owner_name" && column.ColumnKind == WorksheetColumnKind.Single);
            Assert.Contains(schema.Columns, column => column.ApiFieldKey == "region" && column.ColumnKind == WorksheetColumnKind.Single);
            Assert.Contains(schema.Columns, column => column.ApiFieldKey == "priority" && column.ColumnKind == WorksheetColumnKind.Single);
            Assert.Contains(schema.Columns, column => column.ApiFieldKey == "status" && column.ColumnKind == WorksheetColumnKind.Single);
            Assert.Equal(6, schema.Columns.Count(column => column.ColumnKind == WorksheetColumnKind.ActivityProperty));
        }

        [Fact]
        public async Task BatchSaveAcceptsLargeBenchmarkPayload()
        {
            var connector = await CreateConnectorAsync();

            var changes = Enumerable.Range(1, 10000)
                .Select(index => new CellChange
                {
                    RowId = $"benchmark-row-{index:D5}",
                    ApiFieldKey = "status",
                    NewValue = "LoadTestStatus",
                })
                .ToArray();

            connector.BatchSave("large-activity-benchmark", changes);

            var rows = connector.Find(
                "large-activity-benchmark",
                new[] { "benchmark-row-00001", "benchmark-row-05000", "benchmark-row-10000" },
                new[] { "status" });

            Assert.Equal(3, rows.Count);
            Assert.All(
                rows,
                row => Assert.Equal("LoadTestStatus", row["status"]?.ToString()));
        }

        private async Task<CurrentBusinessSystemConnector> CreateConnectorAsync()
        {
            var cookieJar = await fixture.LoginAs("connector_user", "password123").ConfigureAwait(false);
            var httpClient = new HttpClient(new HttpClientHandler
            {
                UseCookies = true,
                CookieContainer = cookieJar,
                SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
            })
            {
                Timeout = TimeSpan.FromSeconds(10),
            };

            return new CurrentBusinessSystemConnector(
                () => new AppSettings { BusinessBaseUrl = fixture.BusinessUrl, ApiKey = string.Empty },
                httpClient);
        }
    }
}
